"""Bounded HTTP retries shared by REST, OWA and upload operations.

Only safe reads retry ambiguous network/5xx failures. An explicit 401 or 429
means the request was rejected and can be retried, without replaying a command.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import random
import fcntl
import os
import weakref
import threading
import time
from urllib.parse import urljoin, urlparse

import httpx

from .exceptions import AmbiguousMutationError, RateLimitError, TokenExpiredError
from .telemetry import record_request

MAX_RETRIES = 3
RETRY_DEADLINE = 90.0
_semaphores: dict[str, threading.BoundedSemaphore] = {}
_guard = threading.Lock()
_refresh_locks = weakref.WeakKeyDictionary()


def retry_after_seconds(value: str | None, *, now: datetime | None = None) -> float | None:
    if not value:
        return None
    try:
        seconds = float(value)
        if not (0 <= seconds < float("inf")):
            return None
        return seconds
    except (TypeError, ValueError):
        try:
            target = parsedate_to_datetime(value)
            if target.tzinfo is None:
                target = target.replace(tzinfo=timezone.utc)
            return max(0, (target - (now or datetime.now(timezone.utc))).total_seconds())
        except (TypeError, ValueError, OverflowError):
            return None


def ensure_next_link(client: httpx.Client, url: str) -> str:
    """Never forward a bearer token to a nextLink on another origin/path."""
    base = str(client.base_url)
    candidate = urljoin(base.rstrip("/") + "/", url)
    expected, actual = urlparse(base), urlparse(candidate)
    if (expected.scheme.lower(), expected.hostname, expected.port) != (actual.scheme.lower(), actual.hostname, actual.port):
        raise ValueError("Refusing pagination URL outside the configured API origin")
    if actual.username or actual.password or not actual.path.startswith(expected.path.rstrip("/") + "/"):
        raise ValueError("Refusing pagination URL outside the configured API path")
    return candidate


@contextmanager
def account_concurrency(account_name: str | None, *, timeout: float = RETRY_DEADLINE):
    """Four in-flight requests per profile across threads and CLI processes."""
    from . import account as account_service
    selected = account_name or "default"
    with _guard:
        semaphore = _semaphores.setdefault(selected, threading.BoundedSemaphore(4))
    acquired = semaphore.acquire(timeout=max(0, timeout))
    if not acquired:
        raise httpx.PoolTimeout("Timed out waiting for the account request limit")
    handle = None
    try:
        directory = account_service.get_account_paths(selected).cache_dir / "http-slots"
        directory.mkdir(parents=True, exist_ok=True)
        end = time.monotonic() + timeout
        while handle is None:
            for slot in range(4):
                candidate = (directory / f"{slot}.lock").open("a+")
                os.chmod(candidate.name, 0o600)
                try:
                    fcntl.flock(candidate.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    handle = candidate
                    break
                except BlockingIOError:
                    candidate.close()
            if handle is None:
                if time.monotonic() >= end:
                    raise httpx.PoolTimeout("Timed out waiting for the account process request limit")
                time.sleep(0.05)
        yield
    finally:
        if handle:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            handle.close()
        semaphore.release()


def request_response(client, method: str, path: str, *, params=None, json=None,
                     headers=None, content=None, refresh=None, account_name=None,
                     retry_safe=None, max_retries=None, deadline=None, telemetry=None,
                     **kwargs):
    method = method.upper()
    safe = method in {"GET", "HEAD", "OPTIONS"} if retry_safe is None else retry_safe
    limit = MAX_RETRIES if max_retries is None else max_retries
    end = time.monotonic() + (RETRY_DEADLINE if deadline is None else deadline)
    attempts = 0
    refreshed = False
    actual_attempt = 0
    request_kwargs = {k: v for k, v in {"params": params, "json": json, "headers": headers, "content": content, **kwargs}.items() if v is not None}
    while True:
        if actual_attempt and time.monotonic() >= end:
            raise httpx.TimeoutException("HTTP request retry deadline exhausted")
        started = time.monotonic()
        actual_attempt += 1
        sent_authorization = request_kwargs.get("headers", {}).get("Authorization") or getattr(client, "headers", {}).get("Authorization")
        try:
            with account_concurrency(account_name, timeout=end - time.monotonic()):
                response = client.request(method, path, **request_kwargs)
        except httpx.PoolTimeout:
            record_request(time.monotonic() - started, retry=actual_attempt > 1)
            raise
        except httpx.RequestError as exc:
            record_request(time.monotonic() - started, retry=actual_attempt > 1)
            if not safe:
                raise AmbiguousMutationError(
                    "The connection failed during a write. The server may have applied it; "
                    "check the affected item before retrying."
                ) from exc
            if attempts >= limit or time.monotonic() >= end:
                raise
            wait = min(2 ** attempts + random.uniform(0, 0.25), max(0, end - time.monotonic()))
            if wait <= 0:
                raise
            attempts += 1
            time.sleep(wait)
            continue
        record_request(time.monotonic() - started, response=response, retry=actual_attempt > 1)
        if telemetry:
            telemetry({"method": method, "status": response.status_code,
                       "elapsed_ms": round((time.monotonic() - started) * 1000), "attempt": attempts + 1})
        if response.status_code == 401:
            if refresh is None or refreshed:
                raise TokenExpiredError("Token expired. Run: outlook login")
            refreshed = True
            with _guard:
                refresh_lock = _refresh_locks.setdefault(client, threading.Lock())
            with refresh_lock:
                current_authorization = client.headers.get("Authorization")
                if current_authorization and current_authorization != sent_authorization:
                    token = current_authorization.removeprefix("Bearer ")
                else:
                    token = refresh()
                if token:
                    client.headers["Authorization"] = f"Bearer {token}"
                    if "headers" in request_kwargs:
                        request_kwargs["headers"] = dict(request_kwargs["headers"])
                        request_kwargs["headers"]["Authorization"] = f"Bearer {token}"
            continue
        retryable = response.status_code == 429 or (safe and response.status_code in {500, 502, 503, 504})
        if retryable:
            wait = retry_after_seconds(response.headers.get("Retry-After"))
            if wait is None:
                wait = 2 ** attempts + random.uniform(0, 0.25)
            if attempts >= limit or time.monotonic() + wait >= end:
                if response.status_code == 429:
                    raise RateLimitError("API rate limit persisted beyond the bounded retry budget. Retry later.")
                response.raise_for_status()
            attempts += 1
            time.sleep(wait)
            continue
        if not safe and response.status_code >= 500:
            raise AmbiguousMutationError(f"Server returned HTTP {response.status_code} during a write. Check the affected item before retrying.")
        response.raise_for_status()
        return response


def request_json(client, method: str, path: str, **kwargs) -> dict:
    response = request_response(client, method, path, **kwargs)
    if response.status_code in {202, 204} or not response.content:
        return {}
    return response.json()
