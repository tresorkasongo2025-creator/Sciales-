"""
Tests for alert email persistence (AppConfig DB table) and the
POST /decanat/bulletins/dashboard/settings route.

Run with:  pytest tests/test_alert_email_settings.py -v
"""

import os
import sys
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))


# ---------------------------------------------------------------------------
# Helpers – spin up a minimal Flask + in-memory SQLite app from app.py
# ---------------------------------------------------------------------------

def _make_app():
    """Return (flask_app, db, _get_alert_email) using an in-memory SQLite DB."""
    import importlib.util
    import types

    spec = importlib.util.spec_from_file_location(
        "app_module_settings",
        os.path.join(os.path.dirname(os.path.dirname(__file__)), "app.py"),
    )
    mod = types.ModuleType("app_module_settings")
    mod.__spec__ = spec

    import flask as _flask
    import flask_sqlalchemy as _fsa

    _app = _flask.Flask(__name__)
    _app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///:memory:"
    _app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    _app.config["SECRET_KEY"] = "test-secret"
    _app.config["TESTING"] = True
    _db = _fsa.SQLAlchemy(_app)

    mod.app = _app
    mod.db = _db

    spec.loader.exec_module(mod)

    # Create all tables in the in-memory DB.
    with _app.app_context():
        _db.create_all()

    return mod


@pytest.fixture(scope="module")
def app_mod():
    return _make_app()


@pytest.fixture()
def ctx(app_mod):
    """Push an app context and yield; teardown cleans AppConfig rows."""
    with app_mod.app.app_context():
        yield app_mod
        # Clean AppConfig between tests
        app_mod.AppConfig.query.delete()
        app_mod.db.session.commit()


@pytest.fixture()
def client(app_mod):
    """Flask test client with a pushed app context; yields (test_client, mod)."""
    with app_mod.app.app_context():
        yield app_mod.app.test_client(), app_mod
        app_mod.AppConfig.query.delete()
        app_mod.db.session.commit()


# ---------------------------------------------------------------------------
# _get_alert_email() – DB persistence
# ---------------------------------------------------------------------------

class TestGetAlertEmail:
    def test_returns_db_value_when_row_exists(self, ctx):
        """_get_alert_email() must return the value stored in AppConfig."""
        mod = ctx
        row = mod.AppConfig(key="alert_email", value="decanat@unilu.ac.cd")
        mod.db.session.add(row)
        mod.db.session.commit()

        result = mod._get_alert_email()
        assert result == "decanat@unilu.ac.cd"

    def test_strips_whitespace_from_db_value(self, ctx):
        mod = ctx
        row = mod.AppConfig(key="alert_email", value="  spaces@example.com  ")
        mod.db.session.add(row)
        mod.db.session.commit()

        result = mod._get_alert_email()
        assert result == "spaces@example.com"

    def test_falls_back_to_env_var_when_no_row(self, ctx, monkeypatch):
        """When AppConfig has no alert_email row, fall back to DECANAT_EMAIL."""
        monkeypatch.setenv("DECANAT_EMAIL", "fallback@unilu.ac.cd")
        result = ctx._get_alert_email()
        assert result == "fallback@unilu.ac.cd"

    def test_falls_back_to_env_var_when_row_is_blank(self, ctx, monkeypatch):
        """An empty DB row must not shadow the env-var fallback."""
        monkeypatch.setenv("DECANAT_EMAIL", "fallback@unilu.ac.cd")
        mod = ctx
        row = mod.AppConfig(key="alert_email", value="")
        mod.db.session.add(row)
        mod.db.session.commit()

        result = mod._get_alert_email()
        assert result == "fallback@unilu.ac.cd"

    def test_returns_empty_string_when_neither_db_nor_env(self, ctx, monkeypatch):
        """When neither DB row nor DECANAT_EMAIL exists, return ''."""
        monkeypatch.delenv("DECANAT_EMAIL", raising=False)
        result = ctx._get_alert_email()
        assert result == ""

    def test_db_takes_priority_over_env_var(self, ctx, monkeypatch):
        """A valid DB value must shadow the DECANAT_EMAIL env var."""
        monkeypatch.setenv("DECANAT_EMAIL", "env@unilu.ac.cd")
        mod = ctx
        row = mod.AppConfig(key="alert_email", value="db@unilu.ac.cd")
        mod.db.session.add(row)
        mod.db.session.commit()

        result = mod._get_alert_email()
        assert result == "db@unilu.ac.cd"

    def test_value_survives_simulated_restart(self, ctx):
        """
        Simulate a server restart: write to DB, clear Python-level caches,
        then read back via a fresh _get_alert_email() call – the value must
        still be there (DB persistence, not in-memory state).
        """
        mod = ctx
        row = mod.AppConfig(key="alert_email", value="persist@unilu.ac.cd")
        mod.db.session.add(row)
        mod.db.session.commit()

        # Expire all identity-map entries to force a real DB round-trip
        mod.db.session.expire_all()

        result = mod._get_alert_email()
        assert result == "persist@unilu.ac.cd"


# ---------------------------------------------------------------------------
# POST /decanat/bulletins/dashboard/settings – route behaviour
# ---------------------------------------------------------------------------

def _logged_in(tc):
    """Set the session flag that decanat_dashboard_settings checks."""
    with tc.session_transaction() as sess:
        sess["decanat_logged_in"] = True
    return tc


class TestSettingsRoute:
    def test_valid_email_is_stored(self, client):
        tc, mod = client
        _logged_in(tc)
        resp = tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={"alert_email": "valid@unilu.ac.cd"},
            follow_redirects=False,
        )
        # Route redirects on success
        assert resp.status_code in (302, 303)

        row = mod.AppConfig.query.filter_by(key="alert_email").first()
        assert row is not None
        assert row.value == "valid@unilu.ac.cd"

    def test_empty_email_clears_db_value(self, client):
        """Submitting an empty email must overwrite any existing DB row with ''."""
        tc, mod = client
        _logged_in(tc)

        # Pre-populate
        tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={"alert_email": "old@unilu.ac.cd"},
            follow_redirects=False,
        )

        # Now clear it
        tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={"alert_email": ""},
            follow_redirects=False,
        )

        row = mod.AppConfig.query.filter_by(key="alert_email").first()
        assert row is not None
        assert row.value == ""

    def test_unauthenticated_request_is_redirected(self, client):
        """Without the session flag the route must redirect to login."""
        tc, mod = client
        resp = tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={"alert_email": "attacker@evil.com"},
            follow_redirects=False,
        )
        assert resp.status_code in (302, 303)
        assert "login" in resp.headers.get("Location", "").lower()

    def test_second_post_updates_existing_row(self, client):
        """Posting twice must update the same row, not create a duplicate."""
        tc, mod = client
        _logged_in(tc)

        tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={"alert_email": "first@unilu.ac.cd"},
            follow_redirects=False,
        )
        tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={"alert_email": "second@unilu.ac.cd"},
            follow_redirects=False,
        )

        rows = mod.AppConfig.query.filter_by(key="alert_email").all()
        assert len(rows) == 1
        assert rows[0].value == "second@unilu.ac.cd"


# ---------------------------------------------------------------------------
# SMTP fields – settings route
# ---------------------------------------------------------------------------

class TestSmtpSettingsRoute:
    """Verify that SMTP fields are written/preserved correctly by the settings route."""

    def test_smtp_fields_are_written_to_appconfig(self, client):
        """POSTing SMTP host/port/user/password must create the corresponding AppConfig rows."""
        tc, mod = client
        _logged_in(tc)

        resp = tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={
                "alert_email": "admin@unilu.ac.cd",
                "smtp_host": "smtp.example.com",
                "smtp_port": "587",
                "smtp_user": "user@example.com",
                "smtp_password": "s3cr3t",
                "smtp_mode": "starttls",
            },
            follow_redirects=False,
        )
        assert resp.status_code in (302, 303)

        host_row = mod.AppConfig.query.filter_by(key="smtp_host").first()
        port_row = mod.AppConfig.query.filter_by(key="smtp_port").first()
        user_row = mod.AppConfig.query.filter_by(key="smtp_user").first()
        pass_row = mod.AppConfig.query.filter_by(key="smtp_password").first()

        assert host_row is not None and host_row.value == "smtp.example.com"
        assert port_row is not None and port_row.value == "587"
        assert user_row is not None and user_row.value == "user@example.com"
        # Password must be stored encrypted (not plain-text)
        assert pass_row is not None
        assert pass_row.value != "s3cr3t", "Password must not be stored in plain text"
        assert pass_row.value.startswith("enc1:"), "Stored password must carry the enc1: prefix"

    def test_blank_smtp_password_does_not_overwrite_existing(self, client):
        """Submitting an empty smtp_password must leave the existing DB password untouched."""
        tc, mod = client
        _logged_in(tc)

        # First POST: set an initial password
        tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={
                "alert_email": "",
                "smtp_host": "smtp.example.com",
                "smtp_port": "587",
                "smtp_user": "user@example.com",
                "smtp_password": "original_secret",
                "smtp_mode": "starttls",
            },
            follow_redirects=False,
        )

        original_stored = mod.AppConfig.query.filter_by(key="smtp_password").first().value

        # Second POST: blank password – must not overwrite
        tc.post(
            "/decanat/bulletins/dashboard/settings",
            data={
                "alert_email": "",
                "smtp_host": "smtp.example.com",
                "smtp_port": "587",
                "smtp_user": "user@example.com",
                "smtp_password": "",   # <-- blank
                "smtp_mode": "starttls",
            },
            follow_redirects=False,
        )

        after_stored = mod.AppConfig.query.filter_by(key="smtp_password").first().value
        assert after_stored == original_stored, (
            "Blank smtp_password submission must not overwrite the existing encrypted password"
        )


# ---------------------------------------------------------------------------
# _get_smtp_config() – env-var fallback
# ---------------------------------------------------------------------------

class TestGetSmtpConfig:
    """Verify priority rules inside _get_smtp_config()."""

    def test_falls_back_to_env_vars_when_no_db_rows(self, ctx, monkeypatch):
        """When AppConfig has no SMTP rows, _get_smtp_config() must use env vars."""
        monkeypatch.setenv("SMTP_HOST", "mail.envhost.com")
        monkeypatch.setenv("SMTP_PORT", "465")
        monkeypatch.setenv("SMTP_USER", "envuser@example.com")
        monkeypatch.setenv("SMTP_PASSWORD", "envpassword")

        mod = ctx
        host, port, user, passw, mode = mod._get_smtp_config()

        assert host == "mail.envhost.com"
        assert port == 465          # must be cast to int
        assert user == "envuser@example.com"
        assert passw == "envpassword"

    def test_db_values_take_priority_over_env_vars(self, ctx, monkeypatch):
        """When AppConfig has SMTP rows, they must shadow the env vars."""
        monkeypatch.setenv("SMTP_HOST", "env.host.com")
        monkeypatch.setenv("SMTP_USER", "envuser@example.com")

        mod = ctx
        # Store DB values
        for key, val in [
            ("smtp_host", "db.host.com"),
            ("smtp_port", "2525"),
            ("smtp_user", "dbuser@example.com"),
        ]:
            mod.db.session.add(mod.AppConfig(key=key, value=val))
        mod.db.session.commit()

        host, port, user, _passw, _mode = mod._get_smtp_config()

        assert host == "db.host.com", "DB smtp_host must shadow SMTP_HOST env var"
        assert port == 2525
        assert user == "dbuser@example.com", "DB smtp_user must shadow SMTP_USER env var"

    def test_default_port_is_587_when_env_absent(self, ctx, monkeypatch):
        """When neither DB nor env var provides a port, default must be 587."""
        monkeypatch.delenv("SMTP_PORT", raising=False)

        _host, port, _user, _passw, _mode = ctx._get_smtp_config()
        assert port == 587
