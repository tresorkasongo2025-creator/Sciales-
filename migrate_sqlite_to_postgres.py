"""Importe la base SQLite locale dans PostgreSQL sans supprimer de données.

Usage:
    python scripts/migrate_sqlite_to_postgres.py

Le script doit être lancé dans le workspace après que DATABASE_URL pointe vers
la base PostgreSQL gérée par Replit. Il utilise SQLite comme source de lecture
et PostgreSQL comme destination. Les lignes sont insérées ou mises à jour par
identifiant primaire ; aucune table n'est supprimée et aucun TRUNCATE/DELETE
n'est exécuté.
"""

from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path

from sqlalchemy import create_engine, text


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "instance" / "esciales_unilu.db"

# Tables ordered so their foreign-key parents are imported first.
TABLES = [
    "app_config",
    "page_contents",
    "etudiants",
    "professeurs",
    "cours",
    "bulletin_sessions",
    "bulletin_data",
    "recus_paiement",
    "presences",
    "horaires",
    "actualites",
    "liste_identifiants",
    "scan_logs",
    "paiement_audits",
    "recours",
]


def quote_identifier(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def main() -> int:
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url.startswith(("postgresql://", "postgres://")):
        print("DATABASE_URL PostgreSQL est absent ou invalide.", file=sys.stderr)
        return 2
    if not SOURCE.exists():
        print(f"Base source introuvable: {SOURCE}", file=sys.stderr)
        return 2

    source = sqlite3.connect(str(SOURCE))
    source.row_factory = sqlite3.Row
    target = create_engine(database_url, pool_pre_ping=True)
    totals: dict[str, int] = {}

    try:
        with target.begin() as connection:
            # Confirm the destination has the expected managed PostgreSQL
            # schema before writing anything.
            existing = {
                row[0]
                for row in connection.execute(
                    text(
                        "SELECT table_name FROM information_schema.tables "
                        "WHERE table_schema = 'public'"
                    )
                )
            }
            missing = [table for table in TABLES if table not in existing]
            if missing:
                raise RuntimeError(
                    "Tables PostgreSQL manquantes: " + ", ".join(missing)
                )

            source_rows = {
                table: source.execute(
                    f"SELECT * FROM {quote_identifier(table)}"
                ).fetchall()
                for table in TABLES
            }

            # The old SQLite database contains audit/receipt rows that point
            # to bulletins deleted before this migration. Keep those
            # relationships valid instead of silently dropping the history.
            # These technical placeholder bulletins are visibly marked and
            # contain no invented student or payment data.
            bulletin_rows = source_rows["bulletin_data"]
            bulletin_columns = list(bulletin_rows[0].keys()) if bulletin_rows else [
                row[1]
                for row in connection.execute(
                    text(
                        "SELECT ordinal_position, column_name "
                        "FROM information_schema.columns "
                        "WHERE table_schema='public' AND table_name='bulletin_data' "
                        "ORDER BY ordinal_position"
                    )
                )
            ]
            existing_bulletin_ids = {row["id"] for row in bulletin_rows}
            referenced_bulletin_ids = set()
            for table, column in (
                ("recus_paiement", "bulletin_id"),
                ("recus_paiement", "tentative_bulletin_id"),
                ("paiement_audits", "bulletin_id"),
            ):
                referenced_bulletin_ids.update(
                    row[column]
                    for row in source_rows[table]
                    if row[column] is not None
                )
            missing_bulletin_ids = sorted(
                referenced_bulletin_ids - existing_bulletin_ids
            )
            for bulletin_id in missing_bulletin_ids:
                placeholder = {column: None for column in bulletin_columns}
                placeholder.update(
                    {
                        "id": bulletin_id,
                        "matricule": f"ARCHIVE-BUL-{bulletin_id}",
                        "nom": f"[Historique] Bulletin manquant #{bulletin_id}",
                        "data_json": (
                            '{{"_migration_placeholder": true, '
                            f'"original_bulletin_id": {bulletin_id}}}'
                        ),
                        "numero_bulletin": f"ARCHIVE-BUL-{bulletin_id}",
                        "paye": False,
                        "montant_paye": 0,
                        "nb_telechargements": 0,
                    }
                )
                source_rows["bulletin_data"].append(placeholder)
            if missing_bulletin_ids:
                print(
                    "Fiches techniques conservées pour bulletins historiques: "
                    + ", ".join(map(str, missing_bulletin_ids))
                )

            for table in TABLES:
                rows = source_rows[table]
                if not rows:
                    totals[table] = 0
                    continue

                columns = list(rows[0].keys())
                column_types = {
                    row[0]: row[1]
                    for row in connection.execute(
                        text(
                            "SELECT column_name, data_type "
                            "FROM information_schema.columns "
                            "WHERE table_schema = 'public' AND table_name = :table"
                        ),
                        {"table": table},
                    )
                }
                quoted_table = quote_identifier(table)
                quoted_columns = ", ".join(
                    quote_identifier(column) for column in columns
                )
                placeholders = ", ".join(
                    f":v_{index}" for index in range(len(columns))
                )
                updates = ", ".join(
                    f"{quote_identifier(column)} = EXCLUDED.{quote_identifier(column)}"
                    for column in columns
                    if column != "id"
                )
                if not updates:
                    updates = '"id" = EXCLUDED."id"'

                statement = text(
                    f"INSERT INTO {quoted_table} ({quoted_columns}) "
                    f"VALUES ({placeholders}) "
                    f'ON CONFLICT ("id") DO UPDATE SET {updates}'
                )

                values = [
                    {
                        f"v_{index}": (
                            None
                            if row[column] is None
                            else bool(row[column])
                            if column_types.get(column) == "boolean"
                            else row[column]
                        )
                        for index, column in enumerate(columns)
                    }
                    for row in rows
                ]
                connection.execute(statement, values)
                totals[table] = len(rows)

            # Explicit IDs do not advance PostgreSQL sequences. Advance every
            # serial sequence so the next new record cannot collide with an
            # imported primary key.
            sequence_rows = connection.execute(
                text(
                    "SELECT table_name, column_name "
                    "FROM information_schema.columns "
                    "WHERE table_schema = 'public' "
                    "AND column_default LIKE 'nextval(%'"
                )
            ).fetchall()
            for table, column in sequence_rows:
                max_id = connection.execute(
                    text(
                        f"SELECT MAX({quote_identifier(column)}) "
                        f"FROM {quote_identifier(table)}"
                    )
                ).scalar()
                if max_id is not None:
                    sequence = connection.execute(
                        text(
                            "SELECT pg_get_serial_sequence(:table_name, :column_name)"
                        ),
                        {
                            "table_name": f"public.{table}",
                            "column_name": column,
                        },
                    ).scalar()
                if max_id is not None and sequence:
                    connection.execute(
                        text(
                            "SELECT setval(CAST(:sequence_name AS regclass), "
                            ":max_id, true)"
                        ),
                        {"sequence_name": sequence, "max_id": max_id},
                    )
    finally:
        source.close()
        target.dispose()

    print("Migration terminée sans suppression.")
    for table, count in totals.items():
        print(f"{table}: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())