"""Tests for exceptions.py — hierarchy and error code mapping."""

from __future__ import annotations

from outlook_cli.exceptions import (
    AuthRequiredError,
    OutlookCliError,
    RateLimitError,
    ResourceNotFoundError,
    TokenExpiredError,
    error_code_for_exception,
)


class TestExceptionHierarchy:
    def test_all_inherit_from_base(self):
        for cls in [TokenExpiredError, RateLimitError, ResourceNotFoundError, AuthRequiredError]:
            assert issubclass(cls, OutlookCliError)
            assert issubclass(cls, Exception)

    def test_base_catches_all(self):
        for cls in [TokenExpiredError, RateLimitError, ResourceNotFoundError, AuthRequiredError]:
            try:
                raise cls("test")
            except OutlookCliError:
                pass  # should be caught

    def test_siblings_dont_catch_each_other(self):
        try:
            raise TokenExpiredError("test")
        except ResourceNotFoundError:
            assert False, "TokenExpiredError should NOT be caught by ResourceNotFoundError"
        except TokenExpiredError:
            pass


class TestErrorCodeMapping:
    def test_known_exceptions(self):
        assert error_code_for_exception(TokenExpiredError("x")) == "session_expired"
        assert error_code_for_exception(RateLimitError("x")) == "rate_limited"
        assert error_code_for_exception(ResourceNotFoundError("x")) == "not_found"
        assert error_code_for_exception(AuthRequiredError("x")) == "not_authenticated"

    def test_unknown_exception(self):
        assert error_code_for_exception(ValueError("x")) == "unknown_error"
        assert error_code_for_exception(RuntimeError("x")) == "unknown_error"
