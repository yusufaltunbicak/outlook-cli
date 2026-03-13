from __future__ import annotations

import json
import os
import stat
import time
from base64 import urlsafe_b64decode
from pathlib import Path

import httpx

from .constants import BASE_URL, BROWSER_STATE_FILE, CACHE_DIR, OWA_URL, TOKEN_FILE, USER_AGENT


def get_token() -> str:
    """Return a valid bearer token, from env, cache, or interactive login."""
    # 1. Environment variable
    env_token = os.environ.get("OUTLOOK_TOKEN")
    if env_token:
        return env_token

    # 2. Cached token
    cached = _load_cached_token()
    if cached:
        return cached

    # 3. Interactive login
    return login()


def login(force: bool = False, debug: bool = False) -> str:
    """Launch Playwright browser to capture bearer token from OWA."""
    from playwright.sync_api import sync_playwright

    CACHE_DIR.mkdir(parents=True, exist_ok=True)

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
        launch_args: dict = {}
        if BROWSER_STATE_FILE.exists() and not force:
            launch_args["storage_state"] = str(BROWSER_STATE_FILE)

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent=USER_AGENT,
            **launch_args,
        )

        # Listen on CONTEXT level to catch all frames/workers
        context.on("request", _intercept_request)

        page = context.new_page()

        print("Opening Outlook... Log in and wait for your inbox to load.")
        print("The browser will close automatically once the token is captured.")
        page.goto(OWA_URL, wait_until="domcontentloaded")

        # Poll: wait up to 120s for a token to appear
        deadline = time.time() + 120
        while not captured_token and time.time() < deadline:
            try:
                page.wait_for_timeout(2000)
            except Exception:
                # Browser was closed by user
                break

            # After 20s without token, try triggering an API call
            if not captured_token and time.time() > deadline - 95:
                try:
                    page.evaluate("""
                        fetch('/api/v2.0/me', {credentials: 'include'})
                            .catch(() => {});
                    """)
                except Exception:
                    pass

        # Save browser state for future SSO
        try:
            context.storage_state(path=str(BROWSER_STATE_FILE))
            _chmod_600(BROWSER_STATE_FILE)
        except Exception:
            pass

        try:
            browser.close()
        except Exception:
            pass

    if debug and seen_urls:
        print(f"\n  [debug] Total requests with Bearer: {len(seen_urls)}")

    if not captured_token:
        raise RuntimeError(
            "Could not capture bearer token.\n"
            "Make sure you logged in and your inbox fully loaded.\n"
            "Tip: Try 'outlook login --debug' to see request details."
        )

    # Deduplicate and pick the best token
    unique_tokens = list(dict.fromkeys(captured_token))  # preserve order, remove dupes
    token = _pick_best_token(unique_tokens, debug=debug)
    _save_token(token)
    return token


def _pick_best_token(tokens: list[str], debug: bool = False) -> str:
    """Try each token against known endpoints. Prefer one that can read mail."""
    import httpx

    # Group by decoded audience
    candidates: list[tuple[str, str]] = []  # (token, audience)
    for t in tokens:
        aud = _decode_audience(t)
        candidates.append((t, aud))

    if debug:
        for t, aud in candidates:
            print(f"  [debug] Token ({len(t)} chars) audience={aud}")

    # Try each token: first against REST v2, then Graph
    endpoints = [
        ("https://outlook.office.com/api/v2.0/me/messages?$top=1", "REST v2"),
        ("https://outlook.office365.com/api/v2.0/me/messages?$top=1", "REST v2 (365)"),
        ("https://graph.microsoft.com/v1.0/me/messages?$top=1", "Graph"),
    ]

    for token, aud in candidates:
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

    # If no token works for mail, try /me as fallback
    for token, aud in candidates:
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

    # Last resort: return the longest token (likely the OWA one)
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
            f"{BASE_URL}",
            headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
            timeout=10,
        )
        return resp.status_code == 200
    except httpx.HTTPError:
        return False


def _load_cached_token() -> str | None:
    if not TOKEN_FILE.exists():
        return None
    try:
        data = json.loads(TOKEN_FILE.read_text())
        token = data["token"]
        expires_at = data.get("expires_at", 0)
        # Check expiry with 5-minute buffer
        if time.time() > expires_at - 300:
            return None
        return token
    except (json.JSONDecodeError, KeyError):
        return None


def _save_token(token: str) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    expires_at = _decode_exp(token)
    data = {"token": token, "expires_at": expires_at}
    TOKEN_FILE.write_text(json.dumps(data))
    _chmod_600(TOKEN_FILE)


def _decode_exp(token: str) -> float:
    """Extract exp claim from JWT without full verification."""
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return time.time() + 3600  # fallback 1h
        payload = parts[1]
        # Fix padding
        payload += "=" * (4 - len(payload) % 4)
        decoded = json.loads(urlsafe_b64decode(payload))
        return float(decoded.get("exp", time.time() + 3600))
    except Exception:
        return time.time() + 3600


def _chmod_600(path: Path) -> None:
    try:
        path.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass
