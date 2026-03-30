"""Structured exception hierarchy for outlook-cli."""

from __future__ import annotations

import click
import httpx


EXIT_CODE_SUCCESS = 0
EXIT_CODE_FAILURE = 1
EXIT_CODE_USAGE = 2
EXIT_CODE_AUTH_REQUIRED = 4
EXIT_CODE_NOT_FOUND = 5
EXIT_CODE_RATE_LIMITED = 7
EXIT_CODE_RETRYABLE = 8
EXIT_CODE_CONFIG = 10
EXIT_CODE_INTERRUPTED = 130


class OutlookCliError(Exception):
    """Base exception for all outlook-cli errors."""


class TokenExpiredError(OutlookCliError):
    """401 — bearer token expired or revoked."""


class RateLimitError(OutlookCliError):
    """429 — API rate limit hit after max retries."""


class ResourceNotFoundError(OutlookCliError):
    """Folder, calendar, category, signature, or message not found."""


class AuthRequiredError(OutlookCliError):
    """No token available — user must run 'outlook login'."""


class AccountError(OutlookCliError):
    """Account profile resolution, binding, or switching failure."""


def error_code_for_exception(exc: Exception) -> str:
    """Map exception to a structured error code string."""
    if isinstance(exc, httpx.HTTPStatusError):
        status = exc.response.status_code if exc.response is not None else None
        if status == 401:
            return "session_expired"
        if status == 404:
            return "not_found"
        if status == 429:
            return "rate_limited"
        if status is not None and status >= 500:
            return "retryable_error"
    if isinstance(exc, httpx.RequestError):
        return "retryable_error"

    mapping = {
        TokenExpiredError: "session_expired",
        RateLimitError: "rate_limited",
        ResourceNotFoundError: "not_found",
        AuthRequiredError: "not_authenticated",
        AccountError: "account_error",
        click.BadParameter: "invalid_usage",
        click.UsageError: "invalid_usage",
    }
    return mapping.get(type(exc), "unknown_error")


def exit_code_for_exception(exc: Exception) -> int:
    """Map exceptions to stable process exit codes for automation."""
    if isinstance(exc, (click.BadParameter, click.UsageError)):
        return EXIT_CODE_USAGE
    if isinstance(exc, (TokenExpiredError, AuthRequiredError)):
        return EXIT_CODE_AUTH_REQUIRED
    if isinstance(exc, ResourceNotFoundError):
        return EXIT_CODE_NOT_FOUND
    if isinstance(exc, RateLimitError):
        return EXIT_CODE_RATE_LIMITED
    if isinstance(exc, AccountError):
        return EXIT_CODE_CONFIG
    if isinstance(exc, KeyboardInterrupt):
        return EXIT_CODE_INTERRUPTED
    if isinstance(exc, httpx.TimeoutException):
        return EXIT_CODE_RETRYABLE
    if isinstance(exc, httpx.RequestError):
        return EXIT_CODE_RETRYABLE
    if isinstance(exc, httpx.HTTPStatusError):
        status = exc.response.status_code if exc.response is not None else None
        if status == 401:
            return EXIT_CODE_AUTH_REQUIRED
        if status == 404:
            return EXIT_CODE_NOT_FOUND
        if status == 429:
            return EXIT_CODE_RATE_LIMITED
        if status is not None and status >= 500:
            return EXIT_CODE_RETRYABLE
    return EXIT_CODE_FAILURE
