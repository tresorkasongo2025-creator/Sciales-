#!/usr/bin/env python3
"""
migrate_to_postgres.py
======================
One-time (idempotent) migration: copies every row from the local SQLite
development database (esciales_unilu.db) into the Replit-managed PostgreSQL
database used by the production deployment.

Usage
-----
  python migrate_to_postgres.py [--sqlite path/to/esciales_unilu.db]

The script reads PostgreSQL credentials from the environment variables that
Replit injects automatically (PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE).
You can also pass a full DATABASE_URL via $DATABASE_URL.

Idempotency
-----------
Every table is migrated with INSERT … ON CONFLICT (id) DO NOTHING, so the
script is safe to run more than once.  After each table the PostgreSQL
auto-increment sequence is bumped so new inserts won't collide with the
migrated IDs.
"""

import argparse
import os
import sqlite3
import sys
from datetime import datetime

import psycopg2
import psycopg2.extras

# ── CLI ────────────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="Migrate SQLite → PostgreSQL")
parser.add_argument(
    "--sqlite",
    default="esciales_unilu.db",
    help="Path to the SQLite database file (default: esciales_unilu.db)",
)
args = parser.parse_args()

SQLITE_PATH = args.sqlite

# ── Validate SQLite file ───────────────────────────────────────────────────────
if not os.path.exists(SQLITE_PATH):
    print(f"[ERROR] SQLite file not found: {SQLITE_PATH}")
    print("        Copy esciales_unilu.db into this directory and try again.")
    sys.exit(1)

# ── PostgreSQL connection ──────────────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")
if DATABASE_URL:
    pg_conn = psycopg2.connect(DATABASE_URL, sslmode="require")
else:
    pg_conn = psycopg2.connect(
        host=os.environ["PGHOST"],
        port=os.environ.get("PGPORT", "5432"),
        user=os.environ["PGUSER"],
        password=os.environ["PGPASSWORD"],
        dbname=os.environ["PGDATABASE"],
        sslmode="require",
    )

pg_conn.autocommit = False
pg = pg_conn.cursor()

# ── SQLite connection ──────────────────────────────────────────────────────────
sq_conn = sqlite3.connect(SQLITE_PATH)
sq_conn.row_factory = sqlite3.Row
sq = sq_conn.cursor()

# ── Helpers ────────────────────────────────────────────────────────────────────

def sqlite_columns(table: str) -> list[str]:
    """Return column names that actually exist in the SQLite table."""
    sq.execute(f"PRAGMA table_info({table})")
    return [row["name"] for row in sq.fetchall()]


def pg_columns(table: str) -> list[str]:
    """Return column names that exist in the PostgreSQL table."""
    pg.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_name = %s
        ORDER BY ordinal_position
        """,
        (table,),
    )
    return [row[0] for row in pg.fetchall()]


def migrate_table(table: str, *, conflict_col: str = "id") -> int:
    """
    Copy all rows from *table* in SQLite into PostgreSQL.
    Only columns present in **both** databases are transferred (handles
    schema drift between dev and prod).
    Uses INSERT … ON CONFLICT (conflict_col) DO NOTHING for idempotency.
    Returns the number of rows inserted.
    """
    sq_cols = set(sqlite_columns(table))
    pg_cols = set(pg_columns(table))

    if not sq_cols:
        print(f"  [SKIP] {table}: not found in SQLite")
        return 0
    if not pg_cols:
        print(f"  [SKIP] {table}: not found in PostgreSQL (run db.create_all() first)")
        return 0

    common = [c for c in sqlite_columns(table) if c in pg_cols]
    if not common:
        print(f"  [SKIP] {table}: no common columns")
        return 0

    sq.execute(f"SELECT {', '.join(common)} FROM {table}")
    rows = sq.fetchall()
    if not rows:
        print(f"  {table}: 0 rows (empty)")
        return 0

    cols_sql = ", ".join(f'"{c}"' for c in common)
    placeholders = ", ".join(["%s"] * len(common))
    insert_sql = (
        f'INSERT INTO "{table}" ({cols_sql}) VALUES ({placeholders}) '
        f'ON CONFLICT ("{conflict_col}") DO NOTHING'
    )

    inserted = 0
    for row in rows:
        values = []
        for col in common:
            val = row[col]
            # SQLite stores booleans as 0/1 integers; convert for PostgreSQL
            if isinstance(val, int) and col in (
                "paye", "utilise", "publie", "tentative_revue"
            ):
                val = bool(val)
            values.append(val)
        pg.execute(insert_sql, values)
        inserted += pg.rowcount

    # Bump the sequence so future inserts don't collide with migrated IDs
    if "id" in common:
        pg.execute(
            f"""
            SELECT setval(
                pg_get_serial_sequence('{table}', 'id'),
                COALESCE((SELECT MAX(id) FROM "{table}"), 1)
            )
            """
        )

    return inserted


# ── Ensure PostgreSQL tables exist ────────────────────────────────────────────
print("Ensuring PostgreSQL schema is up to date …")
# Import the Flask app just for db.create_all(); we do NOT start the server.
# Set env so the app picks PostgreSQL even outside REPLIT_DEPLOYMENT.
os.environ.setdefault("REPLIT_DEPLOYMENT", "1")

try:
    # Temporarily close our own pg connection to avoid connection-limit issues
    # during app import (SQLAlchemy will open its own pool).
    import importlib
    import sys as _sys

    # We only need create_all; import carefully to avoid side effects.
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from app import app, db  # noqa: E402

    with app.app_context():
        db.create_all()
    print("  Schema OK.\n")
except Exception as exc:
    print(f"  [WARN] Could not auto-create schema via app.py: {exc}")
    print("         Make sure the app has been started at least once so the")
    print("         tables exist in PostgreSQL before running this script.\n")

# ── Migration order (respects foreign-key dependencies) ───────────────────────
#
# Independent tables first, then tables that reference them.
#
TABLES = [
    # no FK deps
    "etudiants",
    "professeurs",
    "horaires",
    "actualites",
    "page_contents",
    "bulletin_sessions",
    "app_config",
    "liste_identifiants",
    # FK: professeurs.id
    "cours",
    # FK: etudiants.id, professeurs.id, cours.id
    "presences",
    # FK: bulletin_sessions.id
    "bulletin_data",
    # FK: bulletin_data.id
    "recus_paiement",
    "paiement_audits",
    # no FK deps (references bulletin_data loosely via code, not FK)
    "scan_logs",
]

print("Starting migration …\n")
total_inserted = 0

for table in TABLES:
    try:
        n = migrate_table(table)
        total_inserted += n
        status = f"{n} rows inserted" if n else "0 rows inserted (already migrated or empty)"
        print(f"  ✓ {table}: {status}")
        pg_conn.commit()
    except Exception as exc:
        pg_conn.rollback()
        print(f"  ✗ {table}: ERROR — {exc}")
        print("    Rolling back this table and continuing …")

print(f"\nDone. Total rows inserted across all tables: {total_inserted}")
print("\nYou can safely re-run this script; it will skip already-migrated rows.")

sq_conn.close()
pg_conn.close()
