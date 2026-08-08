"""
Tests for _parse_grille() — guard against silent regressions after parser updates.

Run with:  pytest tests/test_parse_grille.py -v
"""

import os
import re
import sys
import pytest

# ── Make app.py importable without starting Flask ───────────────────────────
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))


def _get_parse_grille():
    """Import _parse_grille from app.py without triggering Flask startup."""
    import importlib.util, types

    spec = importlib.util.spec_from_file_location(
        "app_module",
        os.path.join(os.path.dirname(os.path.dirname(__file__)), "app.py"),
    )
    mod = types.ModuleType("app_module")
    mod.__spec__ = spec

    # Minimal stubs so module-level Flask/SQLAlchemy setup doesn't crash.
    import flask as _flask
    import flask_sqlalchemy as _fsa
    _app = _flask.Flask(__name__)
    _app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///:memory:"
    _app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    _db = _fsa.SQLAlchemy(_app)
    mod.app = _app
    mod.db  = _db

    spec.loader.exec_module(mod)
    return mod._parse_grille


# ── Fixture paths ────────────────────────────────────────────────────────────
_ROOT         = os.path.dirname(os.path.dirname(__file__))
ASSETS_DIR    = os.path.join(_ROOT, "attached_assets")
FIXTURES_DIR  = os.path.join(os.path.dirname(__file__), "fixtures")

REAL_XLS             = os.path.join(
    ASSETS_DIR,
    "Grille_BAC_1_RI_2ème_session_2024-2025_1784247702326.xls",
)
MERGED_TITLE_XLSX    = os.path.join(FIXTURES_DIR, "grille_merged_title.xlsx")
EXTRA_COLS_XLSX      = os.path.join(FIXTURES_DIR, "grille_extra_student_cols.xlsx")

# Matricule pattern for BAC 1 RI: ###-###-###
_MAT_RE = re.compile(r"^\d{3}-\d{3}-\d{3}$")


# ── Module-level cached parse results (avoids re-parsing per test) ───────────

@pytest.fixture(scope="module")
def parsed_real_xls():
    parse = _get_parse_grille()
    return parse(REAL_XLS, "xls")


@pytest.fixture(scope="module")
def parsed_merged_title():
    parse = _get_parse_grille()
    return parse(MERGED_TITLE_XLSX, "xlsx")


@pytest.fixture(scope="module")
def parsed_extra_cols():
    parse = _get_parse_grille()
    return parse(EXTRA_COLS_XLSX, "xlsx")


# ════════════════════════════════════════════════════════════════════════════
# Test 1 — Real BAC 1 RI XLS file
# ════════════════════════════════════════════════════════════════════════════

class TestRealBac1RI:
    """Parse the real production grille and assert known-good counts."""

    def test_course_count(self, parsed_real_xls):
        courses, _, meta = parsed_real_xls
        assert meta["nb_cours"] == 22, (
            f"Expected 22 courses, got {meta['nb_cours']}"
        )
        assert len(courses) == 22

    def test_student_count(self, parsed_real_xls):
        _, students, _ = parsed_real_xls
        assert len(students) == 213, (
            f"Expected 213 students, got {len(students)}"
        )

    def test_no_empty_nom(self, parsed_real_xls):
        """Every parsed student row must have a non-empty nom."""
        _, students, _ = parsed_real_xls
        empty_noms = [s["num"] for s in students if not s["nom"].strip()]
        assert empty_noms == [], f"Students with empty nom: {empty_noms}"

    def test_matricule_format_for_populated_rows(self, parsed_real_xls):
        """
        Students that carry a standard UNILU matricule (###-###-###) must match
        the pattern.  Rows with empty matricule are allowed (late registrants).
        The real file has 2 rows where the matricule cell holds a stray numeric
        ("8.0") — source-data issue, not a parser bug.  We tolerate ≤ 3 such
        anomalies but fail on more to catch new regressions.
        """
        _, students, _ = parsed_real_xls
        bad = [
            (s["num"], s["matricule"])
            for s in students
            if s["matricule"]
            and not _MAT_RE.match(s["matricule"])
        ]
        assert len(bad) <= 3, (
            f"Too many malformed matricules ({len(bad)}): {bad[:10]}"
        )

    def test_majority_of_students_have_valid_matricule(self, parsed_real_xls):
        """At least 200 of 213 students should have a well-formed matricule."""
        _, students, _ = parsed_real_xls
        valid = [s for s in students if _MAT_RE.match(s["matricule"])]
        assert len(valid) >= 200, (
            f"Only {len(valid)} students with valid matricule format (expected ≥ 200)"
        )

    def test_hdr_row_is_zero(self, parsed_real_xls):
        """Header is in the very first row of this file."""
        _, _, meta = parsed_real_xls
        assert meta["hdr_row"] == 0

    def test_course_start_col(self, parsed_real_xls):
        """Courses begin at column 4 (after N°, Matricule, Nom, Sexe)."""
        _, _, meta = parsed_real_xls
        assert meta["course_start_col"] == 4

    def test_each_student_has_correct_course_count(self, parsed_real_xls):
        courses, students, _ = parsed_real_xls
        nb = len(courses)
        bad = [s["num"] for s in students if len(s["cours"]) != nb]
        assert bad == [], (
            f"Students with wrong nb of cours entries: {bad[:5]}"
        )

    def test_promo_dept_extracted_from_sheet_name(self, parsed_real_xls):
        _, _, meta = parsed_real_xls
        assert "BAC 1 RI" in meta["promo_dept"]


# ════════════════════════════════════════════════════════════════════════════
# Test 2 — XLSX with a merged title row (hdr_row ≠ 0)
# ════════════════════════════════════════════════════════════════════════════

class TestMergedTitleRow:
    """
    Grille where row 0 is a merged decorative title and real column headers
    sit on row 1.  The parser must detect hdr_row == 1.
    """

    def test_hdr_row_not_zero(self, parsed_merged_title):
        _, _, meta = parsed_merged_title
        assert meta["hdr_row"] != 0, (
            f"Parser should detect hdr_row > 0 for merged-title file, "
            f"got hdr_row={meta['hdr_row']}"
        )

    def test_hdr_row_is_one(self, parsed_merged_title):
        _, _, meta = parsed_merged_title
        assert meta["hdr_row"] == 1, (
            f"Expected hdr_row=1, got {meta['hdr_row']}"
        )

    def test_nb_cours_positive(self, parsed_merged_title):
        courses, _, meta = parsed_merged_title
        assert meta["nb_cours"] > 0
        assert len(courses) > 0

    def test_students_parsed(self, parsed_merged_title):
        _, students, _ = parsed_merged_title
        assert len(students) > 0

    def test_no_empty_matricule_for_known_students(self, parsed_merged_title):
        """All fixture students were given a matricule; none should be empty."""
        _, students, _ = parsed_merged_title
        empty = [s["num"] for s in students if not s["matricule"]]
        assert empty == [], f"Unexpected empty matricule(s): {empty}"


# ════════════════════════════════════════════════════════════════════════════
# Test 3 — XLSX where course_start_col > 4  (extra student columns)
# ════════════════════════════════════════════════════════════════════════════

class TestExtraStudentColumns:
    """
    Grille with an additional blank column before the N°/Matricule/Nom/Sexe
    block, pushing courses to start at column 5.  course_start_col must be > 4.
    """

    def test_course_start_col_gt_4(self, parsed_extra_cols):
        _, _, meta = parsed_extra_cols
        assert meta["course_start_col"] > 4, (
            f"Expected course_start_col > 4, got {meta['course_start_col']}"
        )

    def test_nb_cours_positive(self, parsed_extra_cols):
        courses, _, meta = parsed_extra_cols
        assert meta["nb_cours"] > 0
        assert len(courses) > 0

    def test_students_parsed(self, parsed_extra_cols):
        _, students, _ = parsed_extra_cols
        assert len(students) > 0

    def test_no_empty_matricule_for_known_students(self, parsed_extra_cols):
        """All fixture students have a matricule; parser must not lose them."""
        _, students, _ = parsed_extra_cols
        empty = [s["num"] for s in students if not s["matricule"]]
        assert empty == [], f"Unexpected empty matricule(s): {empty}"

    def test_student_names_preserved(self, parsed_extra_cols):
        """Nom column must map to the actual name, not a course header."""
        _, students, _ = parsed_extra_cols
        for s in students:
            # Names in this fixture are plain words — should not start with
            # a digit (which would indicate a course number leaked into nom).
            assert not s["nom"][0].isdigit(), (
                f"Student nom looks like a course cell: {s['nom']!r}"
            )
