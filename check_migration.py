#!/usr/bin/env python3
"""
check_migration.py
==================
Sanity-check tool: compares row counts between the local SQLite database and
the Replit-managed PostgreSQL database, table by table.

Usage
-----
  python check_migration.py [--sqlite path/to/esciales_unilu.db]

Exit codes
----------
  0  — all counts match
  1  — one or more tables have a mismatch (or a connection error occurred)

The script reads PostgreSQL credentials from the environment variables that
Replit injects automatically (PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE).
You can also pass a full DATABASE_URL via $DATABASE_URL.
"""

import argparse
import os
import sqlite3
import sys

import psycopg2

# ── CLI ────────────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(
    description="Compare SQLite vs PostgreSQL row counts after migration"
)
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
try:
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
except Exception as exc:
    print(f"[ERROR] Could not connect to PostgreSQL: {exc}")
    sys.exit(1)

pg_conn.autocommit = True
pg = pg_conn.cursor()

# ── SQLite connection ──────────────────────────────────────────────────────────
sq_conn = sqlite3.connect(SQLITE_PATH)
sq = sq_conn.cursor()

# ── Tables to check (same order as migrate_to_postgres.py) ────────────────────
TABLES = [
    "etudiants",
    "professeurs",
    "horaires",
    "actualites",
    "page_contents",
    "bulletin_sessions",
    "app_config",
    "liste_identifiants",
    "cours",
    "presences",
    "bulletin_data",
    "recus_paiement",
    "paiement_audits",
    "scan_logs",
]


def sqlite_table_exists(table: str) -> bool:
    sq.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
    )
    return sq.fetchone() is not None


def pg_table_exists(table: str) -> bool:
    pg.execute(
        "SELECT 1 FROM information_schema.tables WHERE table_name = %s", (table,)
    )
    return pg.fetchone() is not None


def count_sqlite(table: str) -> int:
    sq.execute(f'SELECT COUNT(*) FROM "{table}"')
    return sq.fetchone()[0]


def count_pg(table: str) -> int:
    pg.execute(f'SELECT COUNT(*) FROM "{table}"')
    return pg.fetchone()[0]


# ── Run checks ────────────────────────────────────────────────────────────────
print(f"\nMigration sanity check")
print(f"  SQLite   : {SQLITE_PATH}")
print(f"  PostgreSQL: {os.environ.get('PGHOST', DATABASE_URL[:40] + '…' if DATABASE_URL else 'unknown')}")
print()
print(f"  {'Table':<30} {'SQLite':>8}  {'PG':>8}  Status")
print(f"  {'-'*30}  {'-'*8}  {'-'*8}  ------")

mismatches = 0
errors = 0

for table in TABLES:
    sq_exists = sqlite_table_exists(table)
    pg_exists = pg_table_exists(table)

    if not sq_exists and not pg_exists:
        print(f"  {'  ' + table:<30} {'—':>8}  {'—':>8}  (not in either DB, skipping)")
        continue
    if not sq_exists:
        print(f"  {'  ' + table:<30} {'—':>8}  {'?':>8}  ⚠ not in SQLite")
        continue
    if not pg_exists:
        print(f"  {'  ' + table:<30} {'?':>8}  {'—':>8}  ⚠ not in PostgreSQL")
        errors += 1
        continue

    try:
        sq_count = count_sqlite(table)
        pg_count = count_pg(table)
    except Exception as exc:
        print(f"  {'  ' + table:<30} {'?':>8}  {'?':>8}  ✗ ERROR: {exc}")
        errors += 1
        continue

    if sq_count == pg_count:
        mark = "✓"
        note = ""
    else:
        mark = "✗"
        diff = pg_count - sq_count
        note = f"  (Δ {diff:+d})"
        mismatches += 1

    print(f"  {mark} {table:<30} {sq_count:>8}  {pg_count:>8}{note}")

print()

if mismatches == 0 and errors == 0:
    print("All counts match. Migration looks complete. ✓")
    exit_code = 0
else:
    if mismatches:
        print(f"⚠  {mismatches} table(s) have count mismatches.")
    if errors:
        print(f"⚠  {errors} table(s) had errors (missing or query failed).")
    print("Re-run migrate_to_postgres.py to fill any gaps, then run this check again.")
    exit_code = 1

sq_conn.close()
pg_conn.close()

sys.exit(exit_code)
