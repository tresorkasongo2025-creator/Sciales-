"""Tests for the independent bulletin access stages."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app import (  # noqa: E402
    _GRID_TYPES,
    _RECOURS_CLAIMS,
    _RECEIPT_PAYMENT_DETAILS,
    _RECEIPT_TO_GRID_TYPE,
    _receipt_pdf_title,
    _receipt_type_from_number,
    _receipt_prefix,
    _type_recu_attendu_pour_bulletin,
)


class _Session:
    def __init__(self, type_grille):
        self.type_grille = type_grille


class _Bulletin:
    def __init__(self, type_grille):
        self.bul_session = _Session(type_grille)


def test_each_consultation_receipt_maps_to_only_one_grid_stage():
    assert _RECEIPT_TO_GRID_TYPE == {
        "bulletin": "initial",
        "resultat_recours": "recours",
        "session_2": "session_2",
        "recours_session_2": "recours_session_2",
    }
    assert set(_RECEIPT_TO_GRID_TYPE.values()) == set(_GRID_TYPES)


def test_stage_receipt_prefixes_do_not_overlap_historical_b_lots():
    assert _receipt_prefix("bulletin", lot="1") == "B1"
    assert _receipt_prefix("bulletin", lot="2") == "B2"
    assert _receipt_prefix("recours") == "RS"
    assert _receipt_prefix("resultat_recours") == "RR"
    assert _receipt_prefix("session_2") == "S2"
    assert _receipt_prefix("recours_session_2_soumission") == "RS2"
    assert _receipt_prefix("recours_session_2") == "R2"


def test_printed_receipt_prefix_is_canonical_for_each_stage():
    assert _receipt_type_from_number("RR-RSR-001-S2-26") == "resultat_recours"
    assert _receipt_type_from_number("RS-SC-001-S2-26") == "recours"
    assert _receipt_type_from_number("S2-SC-001-S2-26") == "session_2"
    assert (
        _receipt_type_from_number("RS2-SC-001-S2-26")
        == "recours_session_2_soumission"
    )
    assert _receipt_type_from_number("R2-SC-001-S2-26") == "recours_session_2"
    assert _receipt_type_from_number("B2-SC-001-S2-26") == "bulletin"
    assert _receipt_type_from_number("not-a-receipt") is None


def test_bulletin_stage_requires_its_matching_receipt():
    assert _type_recu_attendu_pour_bulletin(_Bulletin("initial")) == "bulletin"
    assert _type_recu_attendu_pour_bulletin(_Bulletin("recours")) == "resultat_recours"
    assert _type_recu_attendu_pour_bulletin(_Bulletin("session_2")) == "session_2"
    assert (
        _type_recu_attendu_pour_bulletin(_Bulletin("recours_session_2"))
        == "recours_session_2"
    )


def test_six_receipt_categories_have_the_required_amounts():
    assert _RECEIPT_PAYMENT_DETAILS == {
        "bulletin": ("5000 FC", "Cinq mille Francs congolais"),
        "recours": ("10 USD", "Dix Dollars américains"),
        "resultat_recours": ("5000 FC", "Cinq mille Francs congolais"),
        "session_2": ("5000 FC", "Cinq mille Francs congolais"),
        "recours_session_2_soumission": (
            "10 USD",
            "Dix Dollars américains",
        ),
        "recours_session_2": ("5000 FC", "Cinq mille Francs congolais"),
    }


def test_each_receipt_category_has_its_official_pdf_title():
    assert _receipt_pdf_title("resultat_recours") == (
        "PREUVE DE PAIEMENT DES RÉSULTATS DU RECOURS SESSION 1"
    )
    assert _receipt_pdf_title("recours_session_2") == (
        "PREUVE DE PAIEMENT DES RÉSULTATS DU RECOURS SESSION 2"
    )
    assert _receipt_pdf_title("recours_session_2_soumission") == (
        "PREUVE DE PAIEMENT DE LA SOUMISSION DU RECOURS SESSION 2"
    )
    assert _receipt_pdf_title("recours") == (
        "PREUVE DE PAIEMENT DU RECOURS SESSION 1"
    )


def test_missing_course_grade_claim_requires_course_details():
    claim = next(item for item in _RECOURS_CLAIMS if item[0] == "manque_cote")
    assert claim[1] == "Manque de cote dans les cours ci-après"
    assert claim[2] == "cours_manquants"