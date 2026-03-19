from __future__ import annotations

import json
import os
import stat
import time
from base64 import urlsafe_b64decode
from pathlib import Path
from typing import Any

import httpx

from . import account as account_service
from .constants import BASE_URL, OWA_URL, USER_AGENT
from .exceptions import AccountError, AuthRequiredError, TokenExpiredError


def get_token(account_name: str | None = None) -> str:
    """Return a valid bearer token for the selected account."""
    selected = account_service.resolve_account_name(account_name)

    env_token = os.environ.get("OUTLOOK_TOKEN")
    if env_token:
        _assert_token_matches_account(env_token, selected, source="OUTLOOK_TOKEN")
        return env_token

    cached = _load_cached_token() if account_name is None else _load_cached_token(account_name)
    if cached:
        return cached

    if account_name is None:
        return login()
    return login(account_name=selected)


def login(
    force: bool = False,
    debug: bool = False,
    account_name: str | None = None,
    allow_create: bool = False,
    token: str | None = None,
) -> str:
    """Authenticate and cache a bearer token.

    Args:
        force: Force re-login, ignore saved session
        debug: Show debug info about captured requests
        account_name: Account profile name
        allow_create: Allow creating new account profile
        token: Pre-fetched bearer token (skips browser if provided)

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
        _save_token(token, selected, mailbox_info)
        return token

    # Otherwise, launch browser to capture token
    from playwright.sync_api import sync_playwright

    selected = account_service.resolve_account_name(account_name, allow_missing=allow_create)
    if not allow_create:
        account_service.ensure_account_known(selected)

    paths = account_service.get_account_paths(selected)
    paths.cache_dir.mkdir(parents=True, exist_ok=True)
    if not paths.uses_legacy_default:
        paths.config_dir.mkdir(parents=True, exist_ok=True)

    captured_token: list[str] = []
    seen_urls: list[str] = []

    def _intercept_request(request):
        auth = request.headers.get("authorization", "")
        if auth.lower().startswith("bearer "):
            token = auth.split(" ", 1)[1]
            if debug:
                seen_urls.append(request.url[:120])
                print(f"  [debug] Bearer token in: {request.url[:120]}")
            if len(token) > 100:
                captured_token.append(token)
                if debug:
                    print(f"  [debug] Captured token ({len(token)} chars)")

    with sync_playwright() as p:
        launch_args: dict[str, Any] = {}
        if paths.browser_state_file.exists() and not force:
            launch_args["storage_state"] = str(paths.browser_state_file)

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(user_agent=USER_AGENT, **launch_args)
        context.on("request", _intercept_request)

        page = context.new_page()
        print("Opening Outlook... Log in and wait for your inbox to load.")
        print("The browser will close automatically once the token is captured.")
        page.goto(OWA_URL, wait_until="domcontentloaded")

        deadline = time.time() + 120
        while not captured_token and time.time() < deadline:
            try:
                page.wait_for_timeout(2000)
            except Exception:
                break

            if not captured_token and time.time() > deadline - 95:
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

        try:
            browser.close()
        except Exception:
            pass

    if debug and seen_urls:
        print(f"\n  [debug] Total requests with Bearer: {len(seen_urls)}")

    if not captured_token:
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
    for token in tokens:
        aud = _decode_audience(token)
        candidates.append((token, aud))

    if debug:
        for token, aud in candidates:
            print(f"  [debug] Token ({len(token)} chars) audience={aud}")

    endpoints = [
        ("https://outlook.office.com/api/v2.0/me/messages?$top=1", "REST v2"),
        ("https://outlook.office365.com/api/v2.0/me/messages?$top=1", "REST v2 (365)"),
        ("https://graph.microsoft.com/v1.0/me/messages?$top=1", "Graph"),
    ]

    for token, _aud in candidates:
        for url, label in endpoints:
            try:
                resp = httpx.get(
                    url,
                    headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
                    timeout=10,
                )
                if resp.status_code == 200:
                    if debug:
                        print(f"  [debug] Token works with {label}!")
                    return token
            except httpx.HTTPError:
                continue

    for token, _aud in candidates:
        for base in ("https://outlook.office.com/api/v2.0", "https://graph.microsoft.com/v1.0"):
            try:
                resp = httpx.get(
                    f"{base}/me",
                    headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
                    timeout=10,
                )
                if resp.status_code == 200:
                    if debug:
                        print(f"  [debug] Token works for /me at {base} (no mail access though)")
                    return token
            except httpx.HTTPError:
                continue

    return max(tokens, key=len)


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


def _load_cached_token(account_name: str | None = None) -> str | None:
    selected = account_service.resolve_account_name(account_name)
    token_file = account_service.get_account_paths(selected).token_file
    if not token_file.exists():
        return None

    try:
        data = json.loads(token_file.read_text())
        token = data["token"]
        expires_at = data.get("expires_at", 0)
        if time.time() > expires_at - 300:
            return None
    except (json.JSONDecodeError, KeyError):
        return None

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


def _save_token(token: str, account_name: str | None = None, mailbox_info: dict[str, str] | None = None) -> None:
    selected = account_service.resolve_account_name(account_name)
    token_file = account_service.get_account_paths(selected).token_file
    token_file.parent.mkdir(parents=True, exist_ok=True)
    info = mailbox_info or {}
    data = {
        "token": token,
        "expires_at": _decode_exp(token),
        "mailbox_id": info.get("mailbox_id"),
        "email": info.get("email"),
        "display_name": info.get("display_name"),
    }
    token_file.write_text(json.dumps(data))
    _chmod_600(token_file)


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
