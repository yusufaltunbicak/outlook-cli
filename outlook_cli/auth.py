from __future__ import annotations

import json
import os
import stat
import sys
import time
from base64 import urlsafe_b64decode
from pathlib import Path
from typing import Any

import httpx
import keyring
import keyring.errors

from . import account as account_service
from . import credentials
from .constants import BASE_URL, KEYRING_SERVICE_NAME, OWA_URL, USER_AGENT
from .exceptions import AccountError, AuthRequiredError, TokenExpiredError
from .locking import atomic_write_json, file_lock


def _diagnostic(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


TOKEN_STORAGE_BACKEND = "keyring"
TOKEN_STORAGE_VERSION = 1
HEADLESS_CAPTURE_TIMEOUT = 30


def get_token(account_name: str | None = None, *, allow_interactive: bool = True) -> str:
    """Return a token, serializing browser refreshes within each profile."""
    selected = account_service.resolve_account_name(account_name)
    env_token = os.environ.get("OUTLOOK_TOKEN")
    if env_token:
        _assert_token_matches_account(env_token, selected, source="OUTLOOK_TOKEN")
        return env_token
    cached = _load_cached_token(selected, allow_interactive=allow_interactive)
    if cached:
        return cached
    return refresh_token("", selected, allow_interactive=allow_interactive)


def refresh_token(previous: str, account_name: str, *, allow_interactive: bool = True) -> str:
    if os.environ.get("OUTLOOK_TOKEN"):
        raise AuthRequiredError("OUTLOOK_TOKEN was rejected or expired. Replace it or run outlook login without OUTLOOK_TOKEN.")
    paths = account_service.get_account_paths(account_name)
    with file_lock(paths.cache_dir / "refresh.lock", timeout=650):
        # Another process may have refreshed while this one waited for the lock.
        cached = _load_cached_token(account_name, allow_interactive=allow_interactive)
        if cached and cached != previous:
            return cached
        if not allow_interactive:
            raise AuthRequiredError("Authentication required in --no-input mode. Run: outlook login")
        return login(account_name=account_name)


def login(
    force: bool = False,
    debug: bool = False,
    account_name: str | None = None,
    allow_create: bool = False,
    token: str | None = None,
    allow_interactive: bool = True,
    headless: bool | None = None,
) -> str:
    """Authenticate and cache a bearer token.

    Args:
        force: Force re-login, ignore saved session
        debug: Show debug info about captured requests
        account_name: Account profile name
        allow_create: Allow creating new account profile
        token: Pre-fetched bearer token (skips browser if provided)
        headless: Explicit browser mode, or profile configuration for refresh.

    Returns:
        Valid bearer token
    """
    # If token is provided directly, skip browser and validate it
    if token is not None:
        parts = token.split(".")
        if len(parts) != 3:
            raise ValueError("Invalid token format. Expected JWT with 3 parts.")
        selected = account_service.resolve_account_name(account_name, allow_missing=allow_create)
        if not allow_create:
            account_service.ensure_account_known(selected)
        me = _get_me_for_token(token)
        account_service.assert_mailbox_matches(selected, me)
        mailbox_info = account_service.bind_account(selected, me)
        _save_token(token, selected, mailbox_info, allow_interactive=allow_interactive)
        return token

    if not allow_interactive:
        raise AuthRequiredError("Browser login is disabled in --no-input mode. Run: outlook login")

    # Otherwise, launch browser to capture token
    from playwright.sync_api import sync_playwright

    selected = account_service.resolve_account_name(account_name, allow_missing=allow_create)
    if not allow_create:
        account_service.ensure_account_known(selected)

    paths = account_service.get_account_paths(selected)
    paths.cache_dir.mkdir(parents=True, exist_ok=True)
    if not paths.uses_legacy_default:
        paths.config_dir.mkdir(parents=True, exist_ok=True)

    browser_config = account_service.load_account_config(selected).get("browser", {})
    use_headless = bool(browser_config.get("headless", False)) if headless is None else headless
    timeout = float(browser_config.get("timeout", 120))
    if not 0 < timeout <= 600:
        raise AccountError("browser.timeout must be between 0 and 600 seconds.")
    if use_headless:
        if force or not paths.browser_state_file.exists():
            raise AuthRequiredError("Headless refresh requires a saved Outlook session. Run: outlook login")
        timeout = min(timeout, HEADLESS_CAPTURE_TIMEOUT)

    captured_token: list[str] = []
    seen_urls: list[str] = []

    def _intercept_request(request):
        auth = request.headers.get("authorization", "")
        if auth.lower().startswith("bearer "):
            token = auth.split(" ", 1)[1]
            if debug:
                seen_urls.append(request.url[:120])
                _diagnostic(f"  [debug] Bearer token in: {request.url[:120]}")
            if len(token) > 100:
                captured_token.append(token)
                if debug:
                    _diagnostic(f"  [debug] Captured token ({len(token)} chars)")

    with sync_playwright() as p:
        launch_args: dict[str, Any] = {}
        if paths.browser_state_file.exists() and not force:
            launch_args["storage_state"] = str(paths.browser_state_file)

        browser = p.chromium.launch(headless=use_headless, timeout=timeout * 1000)
        try:
            context = browser.new_context(user_agent=USER_AGENT, **launch_args)
            context.on("request", _intercept_request)

            page = context.new_page()
            if use_headless:
                _diagnostic("Refreshing the saved Outlook session...")
            else:
                _diagnostic("Opening Outlook... Log in and wait for your inbox to load.")
                _diagnostic("The browser will close automatically once the token is captured.")
            page.goto(OWA_URL, wait_until="domcontentloaded", timeout=timeout * 1000)

            started = time.monotonic()
            deadline = started + timeout
            nudge_at = started + (0 if use_headless else min(25, timeout))
            while not captured_token and time.monotonic() < deadline:
                try:
                    page.wait_for_timeout(2000)
                except Exception:
                    break

                if not captured_token and time.monotonic() >= nudge_at:
                    try:
                        page.evaluate(
                            """
                            fetch('/api/v2.0/me', {credentials: 'include'})
                                .catch(() => {});
                            """
                        )
                    except Exception:
                        pass

            try:
                context.storage_state(path=str(paths.browser_state_file))
                _chmod_600(paths.browser_state_file)
            except Exception:
                pass
        finally:
            try:
                browser.close()
            except Exception:
                pass

    if debug and seen_urls:
        _diagnostic(f"\n  [debug] Total requests with Bearer: {len(seen_urls)}")

    if not captured_token:
        if use_headless:
            raise AuthRequiredError("Could not refresh the saved Outlook session headlessly. Run: outlook login")
        raise AuthRequiredError(
            "Could not capture bearer token.\n"
            "Make sure you logged in and your inbox fully loaded.\n"
            "Tip: Try 'outlook login --debug' to see request details."
        )

    unique_tokens = list(dict.fromkeys(captured_token))
    token = _pick_best_token(unique_tokens, debug=debug)
    me = _get_me_for_token(token)
    account_service.assert_mailbox_matches(selected, me)
    mailbox_info = account_service.bind_account(selected, me)
    _save_token(token, selected, mailbox_info)
    return token


def _pick_best_token(tokens: list[str], debug: bool = False) -> str:
    """Try each token against known endpoints. Prefer one that can read mail."""
    candidates: list[tuple[str, str]] = []
    probe_deadline = time.monotonic() + 60
    for token in tokens[:8]:
        aud = _decode_audience(token)
        candidates.append((token, aud))

    if debug:
        for token, aud in candidates:
            _diagnostic(f"  [debug] Token ({len(token)} chars) audience={aud}")

    endpoints = [
        ("https://outlook.office.com/api/v2.0/me/messages?$top=1", "REST v2"),
        ("https://outlook.office365.com/api/v2.0/me/messages?$top=1", "REST v2 (365)"),
    ]

    for token, _aud in candidates:
        for url, label in endpoints:
            if time.monotonic() >= probe_deadline:
                raise AuthRequiredError("Mailbox token verification timed out. Run: outlook login")
            try:
                resp = httpx.get(
                    url,
                    headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
                    timeout=max(0.1, min(10, probe_deadline - time.monotonic())),
                )
                if resp.status_code == 200:
                    if debug:
                        _diagnostic(f"  [debug] Token works with {label}!")
                    return token
            except httpx.HTTPError:
                continue

    raise AuthRequiredError("None of the captured tokens could access the mailbox. Run: outlook login --force")


def _decode_audience(token: str) -> str:
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return "unknown"
        payload = parts[1]
        payload += "=" * (4 - len(payload) % 4)
        decoded = json.loads(urlsafe_b64decode(payload))
        return decoded.get("aud", "unknown")
    except Exception:
        return "unknown"


def verify_token(token: str) -> bool:
    """Check if token is valid by calling /me endpoint."""
    try:
        resp = httpx.get(
            BASE_URL,
            headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
            timeout=10,
        )
        return resp.status_code == 200
    except httpx.HTTPError:
        return False


def _load_cached_token(account_name: str | None = None, *, allow_interactive: bool = True) -> str | None:
    selected = account_service.resolve_account_name(account_name)
    token_file = account_service.get_account_paths(selected).token_file
    if not token_file.exists():
        return None

    try:
        data = json.loads(token_file.read_text())
    except json.JSONDecodeError:
        return None

    if "token" in data:
        if not allow_interactive:
            raise AuthRequiredError("Legacy token storage requires interactive migration. Run: outlook login")
        token = data["token"]
        info = {
            "mailbox_id": data.get("mailbox_id"),
            "email": data.get("email"),
            "display_name": data.get("display_name"),
        }
        _save_token(token, selected, info)
        data = _load_token_metadata(token_file) or {}
    expires_at = data.get("expires_at", 0)
    if time.time() > expires_at - 300:
        return None
    token = _load_token_secret(selected, allow_interactive=allow_interactive)

    cached_mailbox = {
        "mailbox_id": data.get("mailbox_id"),
        "email": data.get("email"),
        "display_name": data.get("display_name"),
    }
    if cached_mailbox["mailbox_id"] or cached_mailbox["email"]:
        account_service.assert_mailbox_matches(selected, cached_mailbox)
    else:
        _assert_token_matches_account(token, selected, source=str(token_file))

    return token


def _save_token(token: str, account_name: str | None = None, mailbox_info: dict[str, str] | None = None, *, allow_interactive: bool = True) -> None:
    selected = account_service.resolve_account_name(account_name)
    token_file = account_service.get_account_paths(selected).token_file
    token_file.parent.mkdir(parents=True, exist_ok=True)
    info = mailbox_info or {}
    _store_token_secret(selected, token, allow_interactive=allow_interactive)
    data = {
        "storage_backend": TOKEN_STORAGE_BACKEND,
        "storage_version": TOKEN_STORAGE_VERSION,
        "expires_at": _decode_exp(token),
        "mailbox_id": info.get("mailbox_id"),
        "email": info.get("email"),
        "display_name": info.get("display_name"),
    }
    atomic_write_json(token_file, data)
    _chmod_600(token_file)


def delete_stored_token(account_name: str | None = None, *, allow_interactive: bool = True) -> None:
    selected = account_service.resolve_account_name(account_name, allow_missing=True)
    try:
        credentials.delete_password(KEYRING_SERVICE_NAME, _keyring_username(selected), allow_interactive=allow_interactive)
    except keyring.errors.PasswordDeleteError:
        pass
    except keyring.errors.KeyringError as exc:
        raise AccountError(f"Could not delete stored token for account '{selected}': {exc}") from exc


def _load_token_metadata(token_file: Path) -> dict[str, Any] | None:
    if not token_file.exists():
        return None
    try:
        return json.loads(token_file.read_text())
    except json.JSONDecodeError:
        return None


def _keyring_username(account_name: str) -> str:
    return f"token:{account_name}"


def _store_token_secret(account_name: str, token: str, *, allow_interactive: bool = True) -> None:
    try:
        credentials.set_password(KEYRING_SERVICE_NAME, _keyring_username(account_name), token, allow_interactive=allow_interactive)
    except keyring.errors.KeyringError as exc:
        raise AccountError(
            f"Could not store token securely for account '{account_name}'. Check keyring availability."
        ) from exc


def _load_token_secret(account_name: str, *, allow_interactive: bool = True) -> str:
    try:
        token = credentials.get_password(KEYRING_SERVICE_NAME, _keyring_username(account_name), allow_interactive=allow_interactive)
    except keyring.errors.KeyringError as exc:
        raise AccountError(
            f"Could not read stored token for account '{account_name}'. Check keyring availability."
        ) from exc
    if not token:
        raise AccountError(
            f"Stored token for account '{account_name}' was not found in the keyring. Run: outlook login"
        )
    return token


def _decode_exp(token: str) -> float:
    """Extract exp claim from JWT without full verification."""
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return time.time() + 3600
        payload = parts[1]
        payload += "=" * (4 - len(payload) % 4)
        decoded = json.loads(urlsafe_b64decode(payload))
        return float(decoded.get("exp", time.time() + 3600))
    except Exception:
        return time.time() + 3600


def _get_me_for_token(token: str) -> dict[str, Any]:
    try:
        resp = httpx.get(
            BASE_URL,
            headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
            timeout=10,
        )
    except httpx.HTTPError as exc:
        raise AccountError(f"Could not verify mailbox for the selected account: {exc}") from exc

    if resp.status_code == 401:
        raise TokenExpiredError("Token expired. Run: outlook login")
    if resp.status_code != 200:
        raise AccountError(
            f"Could not verify mailbox for the selected account (HTTP {resp.status_code})."
        )
    return resp.json()


def _assert_token_matches_account(token: str, account_name: str, source: str) -> dict[str, str]:
    bound = account_service.get_account(account_name)
    if not bound.get("mailbox_id"):
        return {}

    me = _get_me_for_token(token)
    try:
        return account_service.assert_mailbox_matches(account_name, me)
    except AccountError as exc:
        raise AccountError(f"{source} belongs to the wrong mailbox. {exc}") from exc


def _chmod_600(path: Path) -> None:
    try:
        path.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass
