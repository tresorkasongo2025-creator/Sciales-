"""Tests for conservative automatic matricule attribution."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app import app, _preparer_attributions_matricules


class _Identifier:
    def __init__(self, nom, matricule):
        self.nom = nom
        self.nom_norm = nom.upper()
        self.matricule = matricule


def test_matching_requires_review_for_ambiguous_missing_and_duplicate_rows():
    rows = [
        _Identifier("KABILA KASONGO JEAN", "MAT-001"),
        _Identifier("KABILA KASONGO JEAN", "MAT-002"),
        _Identifier("MUKENDI KALALA ROSE", "MAT-003"),
    ]
    students = [
        {"nom": "MUKENDI KALALA ROSE", "matricule": "GRID-01", "sexe": "F"},
        {"nom": "KABILA KASONGO JEAN", "matricule": "GRID-02", "sexe": "M"},
        {"nom": "INCONNU TEST", "matricule": "GRID-03", "sexe": "M"},
        {"nom": "MUKENDI KALALA ROSE", "matricule": "GRID-04", "sexe": "F"},
    ]

    with app.app_context():
        result = _preparer_attributions_matricules(students, rows)

    assert result[0]["statut"] == "doublon"
    assert result[1]["statut"] == "ambigu"
    assert result[2]["statut"] == "introuvable"
    assert result[3]["statut"] == "doublon"
    assert all(item["matricule"] == "" for item in result)


def test_exact_match_replaces_grid_matricule_only_with_official_value():
    rows = [_Identifier("KALALA ROSE", "OFFICIAL-001")]
    students = [{"nom": "KALALA ROSE", "matricule": "WRONG-GRID-VALUE", "sexe": "F"}]

    with app.app_context():
        result = _preparer_attributions_matricules(students, rows)

    assert result == [{
        "index": 0,
        "nom": "KALALA ROSE",
        "matricule_grille": "WRONG-GRID-VALUE",
        "statut": "confirme",
        "type_correspondance": "nom complet",
        "matricule": "OFFICIAL-001",
        "nom_liste": "KALALA ROSE",
        "candidats": [{"matricule": "OFFICIAL-001", "nom": "KALALA ROSE"}],
    }]
