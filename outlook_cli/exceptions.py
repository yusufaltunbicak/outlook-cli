"""Structured exception hierarchy for outlook-cli."""

from __future__ import annotations


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


def error_code_for_exception(exc: Exception) -> str:
    """Map exception to a structured error code string."""
    mapping = {
        TokenExpiredError: "session_expired",
        RateLimitError: "rate_limited",
        ResourceNotFoundError: "not_found",
        AuthRequiredError: "not_authenticated",
    }
    return mapping.get(type(exc), "unknown_error")
