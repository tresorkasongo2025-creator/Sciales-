#!/bin/bash
set -e

echo "=== Post-merge setup ==="

# Installer les dépendances Python
echo "Installation des dépendances..."
pip install -r requirements.txt --quiet

# Appliquer les migrations SQLite manquantes (idempotent)
echo "Vérification des migrations..."
python3 - <<'PYEOF'
import sqlite3, os

db_path = os.path.join(os.path.dirname(__file__), '..', 'instance', 'esciales_unilu.db')
if not os.path.exists(db_path):
    print("Base de données introuvable — sera créée au démarrage.")
else:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Colonnes à ajouter si absentes
    migrations = [
        ("bulletin_sessions", "departement",   "TEXT DEFAULT ''"),
        ("bulletin_sessions", "texte_intro",   "TEXT DEFAULT ''"),
        ("recus_paiement",    "semestre",       "TEXT DEFAULT ''"),
        ("recus_paiement",    "montant",        "INTEGER DEFAULT 0"),
        ("recus_paiement",    "montant_lettres","TEXT DEFAULT ''"),
        ("recus_paiement",    "motif",          "TEXT DEFAULT ''"),
        ("recus_paiement",    "annee_complete", "TEXT DEFAULT ''"),
    ]

    for table, column, col_def in migrations:
        try:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_def}")
            conn.commit()
            print(f"  + {table}.{column} ajouté")
        except sqlite3.OperationalError as e:
            if "duplicate column" in str(e).lower():
                pass   # déjà présent
            else:
                print(f"  ! {table}.{column} : {e}")

    conn.close()
    print("Migrations OK.")
PYEOF

echo "=== Setup terminé ==="
