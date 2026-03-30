"""Tests for auth.py token selection, cache handling, and browser login flow."""

from __future__ import annotations

import json
import sys
import types
from base64 import urlsafe_b64encode
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from outlook_cli import auth
from outlook_cli.account import AccountPaths
from outlook_cli.exceptions import AccountError, AuthRequiredError


def _jwt(payload: dict) -> str:
    header = urlsafe_b64encode(json.dumps({"alg": "none"}).encode()).decode().rstrip("=")
    body = urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")
    return f"{header}.{body}.signature"


def _account_paths(tmp_path: Path) -> AccountPaths:
    return AccountPaths(
        name="default",
        cache_dir=tmp_path,
        config_dir=tmp_path,
        token_file=tmp_path / "token.json",
        browser_state_file=tmp_path / "browser-state.json",
        id_map_file=tmp_path / "id_map.json",
        scheduled_file=tmp_path / "scheduled.json",
        signatures_dir=tmp_path / "signatures",
        profile_config_file=tmp_path / "config.yaml",
    )


def _patch_account(monkeypatch, tmp_path: Path, *, bound: dict | None = None):
    paths = _account_paths(tmp_path)
    monkeypatch.setattr(
        auth.account_service,
        "resolve_account_name",
        lambda account_name=None, allow_missing=False: account_name or "default",
    )
    monkeypatch.setattr(auth.account_service, "get_account_paths", lambda account_name: paths)
    monkeypatch.setattr(auth.account_service, "get_account", lambda account_name: {"name": account_name, **(bound or {})})
    monkeypatch.setattr(
        auth.account_service,
        "assert_mailbox_matches",
        lambda account_name, me: {
            "mailbox_id": me.get("Id") or me.get("mailbox_id") or (me.get("EmailAddress") or me.get("email", "")).lower(),
            "email": me.get("EmailAddress") or me.get("email"),
            "display_name": me.get("DisplayName") or me.get("display_name"),
        },
    )
    monkeypatch.setattr(
        auth.account_service,
        "bind_account",
        lambda account_name, me: {
            "mailbox_id": me.get("Id") or (me.get("EmailAddress") or "").lower(),
            "email": me.get("EmailAddress"),
            "display_name": me.get("DisplayName"),
        },
    )
    return paths


def _patch_keyring(monkeypatch):
    store: dict[tuple[str, str], str] = {}

    monkeypatch.setattr(auth.keyring, "set_password", lambda service, username, token: store.__setitem__((service, username), token))
    monkeypatch.setattr(auth.keyring, "get_password", lambda service, username: store.get((service, username)))

    def delete_password(service, username):
        key = (service, username)
        if key not in store:
            raise auth.keyring.errors.PasswordDeleteError("missing")
        del store[key]

    monkeypatch.setattr(auth.keyring, "delete_password", delete_password)
    return store


class _FakePage:
    def __init__(self, context, token: str | None, raise_on_wait: bool):
        self._context = context
        self._token = token
        self._raise_on_wait = raise_on_wait
        self._triggered = False

    def goto(self, *_args, **_kwargs):
        return None

    def wait_for_timeout(self, *_args, **_kwargs):
        if self._raise_on_wait:
            raise Exception("browser closed")
        if self._token and not self._triggered and self._context.callback:
            self._triggered = True
            request = types.SimpleNamespace(
                headers={"authorization": f"Bearer {self._token}"},
                url="https://outlook.office.com/api/v2.0/me/messages?$top=1",
            )
            self._context.callback(request)

    def evaluate(self, *_args, **_kwargs):
        return None


class _FakeContext:
    def __init__(self, token: str | None, raise_on_wait: bool):
        self.token = token
        self.raise_on_wait = raise_on_wait
        self.callback = None
        self.storage_path = None

    def on(self, _event, callback):
        self.callback = callback

    def new_page(self):
        return _FakePage(self, self.token, self.raise_on_wait)

    def storage_state(self, path: str):
        self.storage_path = path
        Path(path).write_text("{}")


class _FakeBrowser:
    def __init__(self, token: str | None, raise_on_wait: bool):
        self.token = token
        self.raise_on_wait = raise_on_wait
        self.context_kwargs = None

    def new_context(self, **kwargs):
        self.context_kwargs = kwargs
        return _FakeContext(self.token, self.raise_on_wait)

    def close(self):
        return None


class _FakePlaywrightCM:
    def __init__(self, token: str | None, raise_on_wait: bool):
        self.browser = _FakeBrowser(token, raise_on_wait)
        self.chromium = types.SimpleNamespace(launch=lambda **_kwargs: self.browser)

    def __enter__(self):
        return types.SimpleNamespace(chromium=self.chromium)

    def __exit__(self, exc_type, exc, tb):
        return False


def _install_fake_playwright(monkeypatch, token: str | None, raise_on_wait: bool = False):
    cm = _FakePlaywrightCM(token, raise_on_wait)
    fake_module = types.SimpleNamespace(sync_playwright=lambda: cm)
    monkeypatch.setitem(sys.modules, "playwright", types.SimpleNamespace(sync_api=fake_module))
    monkeypatch.setitem(sys.modules, "playwright.sync_api", fake_module)
    return cm


def test_get_token_prefers_environment(monkeypatch, tmp_path):
    _patch_account(monkeypatch, tmp_path)
    monkeypatch.setenv("OUTLOOK_TOKEN", "env-token")
    monkeypatch.setattr(auth, "_load_cached_token", lambda: "cached-token")
    monkeypatch.setattr(auth, "login", lambda: "fresh-token")
    monkeypatch.setattr(auth, "_assert_token_matches_account", lambda *args, **kwargs: {})

    assert auth.get_token() == "env-token"


def test_get_token_uses_cache_before_login(monkeypatch, tmp_path):
    _patch_account(monkeypatch, tmp_path)
    monkeypatch.delenv("OUTLOOK_TOKEN", raising=False)
    monkeypatch.setattr(auth, "_load_cached_token", lambda: "cached-token")
    monkeypatch.setattr(auth, "login", lambda: "fresh-token")

    assert auth.get_token() == "cached-token"


def test_get_token_rejects_wrong_mailbox_for_bound_profile(monkeypatch, tmp_path):
    _patch_account(monkeypatch, tmp_path, bound={"mailbox_id": "expected", "email": "user@example.com"})
    monkeypatch.setenv("OUTLOOK_TOKEN", "env-token")
    monkeypatch.setattr(
        auth,
        "_assert_token_matches_account",
        lambda *args, **kwargs: (_ for _ in ()).throw(AccountError("wrong mailbox")),
    )

    with pytest.raises(AccountError, match="wrong mailbox"):
        auth.get_token()


def test_decode_helpers_parse_jwt():
    token = _jwt({"aud": "outlook", "exp": 1890000000})

    assert auth._decode_audience(token) == "outlook"
    assert auth._decode_exp(token) == 1890000000.0


def test_load_cached_token_honors_expiry(monkeypatch, tmp_path):
    paths = _patch_account(monkeypatch, tmp_path)
    store = _patch_keyring(monkeypatch)
    monkeypatch.setattr(auth.time, "time", lambda: 1_000)
    monkeypatch.setattr(auth, "_assert_token_matches_account", lambda *args, **kwargs: {})

    store[(auth.KEYRING_SERVICE_NAME, auth._keyring_username("default"))] = "cached"
    paths.token_file.write_text(json.dumps({"storage_backend": "keyring", "storage_version": 1, "expires_at": 2_000}))
    assert auth._load_cached_token() == "cached"

    paths.token_file.write_text(json.dumps({"storage_backend": "keyring", "storage_version": 1, "expires_at": 1_200}))
    assert auth._load_cached_token() is None

    paths.token_file.write_text("{not-json")
    assert auth._load_cached_token() is None


def test_save_token_writes_expected_payload(monkeypatch, tmp_path):
    paths = _patch_account(monkeypatch, tmp_path)
    store = _patch_keyring(monkeypatch)
    monkeypatch.setattr(auth, "_decode_exp", lambda _token: 1234.0)
    chmod_calls = []
    monkeypatch.setattr(auth, "_chmod_600", lambda path: chmod_calls.append(path))

    auth._save_token("saved-token", mailbox_info={"mailbox_id": "m-1", "email": "u@example.com", "display_name": "User"})

    assert json.loads(paths.token_file.read_text()) == {
        "storage_backend": "keyring",
        "storage_version": 1,
        "expires_at": 1234.0,
        "mailbox_id": "m-1",
        "email": "u@example.com",
        "display_name": "User",
    }
    assert store[(auth.KEYRING_SERVICE_NAME, auth._keyring_username("default"))] == "saved-token"
    assert chmod_calls == [paths.token_file]


def test_load_cached_token_migrates_legacy_plaintext_token(monkeypatch, tmp_path):
    paths = _patch_account(monkeypatch, tmp_path)
    store = _patch_keyring(monkeypatch)
    monkeypatch.setattr(auth.time, "time", lambda: 1_000)
    monkeypatch.setattr(auth, "_assert_token_matches_account", lambda *args, **kwargs: {})

    paths.token_file.write_text(
        json.dumps(
            {
                "token": "legacy-token",
                "expires_at": 2_000,
                "mailbox_id": "mailbox-1",
                "email": "user@example.com",
                "display_name": "User",
            }
        )
    )

    assert auth._load_cached_token() == "legacy-token"
    assert store[(auth.KEYRING_SERVICE_NAME, auth._keyring_username("default"))] == "legacy-token"
    migrated = json.loads(paths.token_file.read_text())
    assert "token" not in migrated
    assert migrated["storage_backend"] == "keyring"
    assert migrated["email"] == "user@example.com"


def test_load_cached_token_requires_keyring_secret(monkeypatch, tmp_path):
    paths = _patch_account(monkeypatch, tmp_path)
    _patch_keyring(monkeypatch)
    paths.token_file.write_text(json.dumps({"storage_backend": "keyring", "storage_version": 1, "expires_at": 2_000}))

    with pytest.raises(AccountError, match="not found in the keyring"):
        auth._load_cached_token()


def test_pick_best_token_prefers_working_mail_endpoint(monkeypatch):
    good = "good-token"
    bad = "bad-token"

    def fake_get(url, headers, timeout):
        token = headers["Authorization"].split(" ", 1)[1]
        status_code = 200 if token == good and "messages" in url else 401
        return types.SimpleNamespace(status_code=status_code)

    monkeypatch.setattr(auth.httpx, "get", fake_get)

    assert auth._pick_best_token([bad, good]) == good


def test_pick_best_token_falls_back_to_longest_token(monkeypatch):
    def fake_get(*_args, **_kwargs):
        raise auth.httpx.HTTPError("network error")

    monkeypatch.setattr(auth.httpx, "get", fake_get)

    assert auth._pick_best_token(["short", "much-longer-token"]) == "much-longer-token"


def test_verify_token_handles_http_error(monkeypatch):
    monkeypatch.setattr(auth.httpx, "get", lambda *_args, **_kwargs: types.SimpleNamespace(status_code=200))
    assert auth.verify_token("token") is True

    def raise_http_error(*_args, **_kwargs):
        raise auth.httpx.HTTPError("boom")

    monkeypatch.setattr(auth.httpx, "get", raise_http_error)
    assert auth.verify_token("token") is False


def test_login_uses_browser_state_and_saves_best_token(monkeypatch, tmp_path):
    token = "x" * 101
    paths = _patch_account(monkeypatch, tmp_path)
    paths.browser_state_file.write_text("{}")
    monkeypatch.setattr(auth, "_chmod_600", lambda _path: None)
    save = MagicMock()
    pick = MagicMock(return_value=token)
    bind = MagicMock(return_value={"mailbox_id": "mailbox-1", "email": "alice@example.com", "display_name": "Alice"})
    monkeypatch.setattr(auth, "_save_token", save)
    monkeypatch.setattr(auth, "_pick_best_token", pick)
    monkeypatch.setattr(auth, "_get_me_for_token", lambda value: {"Id": "mailbox-1", "EmailAddress": "alice@example.com", "DisplayName": "Alice"} if value == token else {})
    monkeypatch.setattr(auth.account_service, "bind_account", bind)
    cm = _install_fake_playwright(monkeypatch, token=token)

    result = auth.login()

    assert result == token
    assert cm.browser.context_kwargs["storage_state"] == str(paths.browser_state_file)
    pick.assert_called_once()
    bind.assert_called_once()
    save.assert_called_once_with(
        token,
        "default",
        {"mailbox_id": "mailbox-1", "email": "alice@example.com", "display_name": "Alice"},
    )


def test_login_raises_when_token_cannot_be_captured(monkeypatch, tmp_path):
    _patch_account(monkeypatch, tmp_path)
    monkeypatch.setattr(auth, "_chmod_600", lambda _path: None)
    _install_fake_playwright(monkeypatch, token=None, raise_on_wait=True)

    with pytest.raises(AuthRequiredError):
        auth.login()
