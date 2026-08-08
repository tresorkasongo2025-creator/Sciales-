"""Tests for payment report date selections."""

from datetime import datetime

from types import SimpleNamespace

from app import (
    _normaliser_filtres_date_paiements,
    _paiement_dans_periode,
    _normaliser_recherche_recu,
    _session_correspond_au_filtre,
)


def test_payment_report_range_includes_both_boundary_days():
    period = _normaliser_filtres_date_paiements(
        "range", date_debut="2026-08-01", date_fin="2026-08-03"
    )

    assert period["label"] == "Du 01/08/2026 au 03/08/2026"
    assert _paiement_dans_periode(datetime(2026, 8, 1, 8), period)
    assert _paiement_dans_periode(datetime(2026, 8, 3, 23, 59), period)
    assert not _paiement_dans_periode(datetime(2026, 8, 4), period)


def test_payment_report_supports_all_exact_and_today_modes():
    all_dates = _normaliser_filtres_date_paiements("all")
    exact = _normaliser_filtres_date_paiements(
        "exact", date_precise="2026-08-02"
    )
    today = _normaliser_filtres_date_paiements(
        "today", today_str="2026-08-03"
    )

    assert all_dates["label"] == "Toutes les dates"
    assert exact["label"] == "02/08/2026"
    assert today["label"] == "03/08/2026"
    assert _paiement_dans_periode(datetime(2026, 8, 2), all_dates)
    assert _paiement_dans_periode(datetime(2026, 8, 2), exact)
    assert not _paiement_dans_periode(datetime(2026, 8, 1), exact)
    assert _paiement_dans_periode(datetime(2026, 8, 3), today)


def test_payment_report_session_filter_separates_speciale_and_recours():
    initial = SimpleNamespace(type_grille="initial")
    session_2 = SimpleNamespace(type_grille="session_2")
    recours = SimpleNamespace(type_grille="recours")
    recours_2 = SimpleNamespace(type_grille="recours_session_2")

    assert _session_correspond_au_filtre(initial, "speciale")
    assert _session_correspond_au_filtre(session_2, "speciale")
    assert not _session_correspond_au_filtre(recours, "speciale")
    assert not _session_correspond_au_filtre(recours_2, "speciale")

    assert not _session_correspond_au_filtre(initial, "recours")
    assert not _session_correspond_au_filtre(session_2, "recours")
    assert _session_correspond_au_filtre(recours, "recours")
    assert _session_correspond_au_filtre(recours_2, "recours")


def test_payment_report_can_filter_second_session_stages_independently():
    initial = SimpleNamespace(type_grille="initial")
    session_2 = SimpleNamespace(type_grille="session_2")
    recours = SimpleNamespace(type_grille="recours")
    recours_2 = SimpleNamespace(type_grille="recours_session_2")

    for bulletin_session, expected in (
        (initial, "initial"),
        (session_2, "session_2"),
        (recours, "recours_1"),
        (recours_2, "recours_session_2"),
    ):
        for stage in (
            "initial", "session_2", "recours_1", "recours_session_2"
        ):
            assert _session_correspond_au_filtre(
                bulletin_session, stage
            ) is (stage == expected)
    assert _session_correspond_au_filtre(recours, "recours_1")
    assert not _session_correspond_au_filtre(recours_2, "recours_1")


def test_receipt_lookup_normalizes_reference_qr_url_and_json():
    assert _normaliser_recherche_recu("b1-ri-001-s2-26") == "B1-RI-001-S2-26"
    assert _normaliser_recherche_recu("/scan/abc123") == "ABC123"
    assert _normaliser_recherche_recu(
        "https://example.test/scan/scan/ABC123"
    ) == "ABC123"
    assert _normaliser_recherche_recu('{"code_qr":"abc123"}') == "ABC123"