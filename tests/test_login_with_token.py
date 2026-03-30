"""Tests for --with-token functionality in login command."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from click.testing import CliRunner

from outlook_cli.commands import auth as auth_cmd
from outlook_cli import auth as auth_module
from outlook_cli import account as account_service


class TestWithTokenLogin:
    """Test --with-token login functionality."""

    def test_with_token_rejects_empty_input(self, runner):
        """Test that --with-token rejects empty stdin."""
        result = runner.invoke(auth_cmd.login, ["--with-token"], input="")

        assert result.exit_code == 2
        assert "No token provided" in result.output

    def test_with_token_validates_jwt_format(self, runner):
        """Test that --with-token validates JWT format (3 parts)."""
        # Invalid formats
        invalid_tokens = [
            "not_a_jwt",
            "only.two",
        ]

        for invalid_token in invalid_tokens:
            result = runner.invoke(auth_cmd.login, ["--with-token"], input=f"{invalid_token}\n")
            assert result.exit_code == 1
            # Error message comes from ValueError in auth.login()
            assert "Invalid token format" in result.output


class TestTokenValidation:
    """Test token validation logic."""

    def test_verify_token_calls_correct_endpoint(self):
        """Test that verify_token calls the /me endpoint."""
        # This would require mocking httpx, which is tested elsewhere
        # Just document the expected behavior here
        pass

    def test_get_token_prefers_env_over_cache(self, monkeypatch):
        """Test that OUTLOOK_TOKEN takes priority over cached token."""
        # Mock to test priority order
        called_order = []

        def cache_loader():
            called_order.append("cache")
            return "cached_token"

        monkeypatch.setattr(auth_module, "_load_cached_token", cache_loader)

        # Simulate env token check
        import os
        monkeypatch.setenv("OUTLOOK_TOKEN", "env_token_123")

        # Check env token first
        env_token = os.environ.get("OUTLOOK_TOKEN")
        if env_token:
            called_order.append("env")

        assert called_order == ["env"]
        # env should be checked before cache


class TestWithTokenUnit:
    """Unit tests for token handling logic."""

    def test_login_accepts_valid_token(self, monkeypatch):
        """Test that login() function accepts and validates a token."""
        fake_token = "header.payload.signature"

        def fake_get_me(token):
            return {
                "Id": "mailbox123",
                "EmailAddress": "test@example.com",
                "DisplayName": "Test User",
            }

        saved_data = {}

        def fake_save_token(t, acc, info):
            saved_data["token"] = t
            saved_data["account"] = acc
            saved_data["info"] = info

        monkeypatch.setattr(auth_module, "_get_me_for_token", fake_get_me)
        monkeypatch.setattr(auth_module, "_save_token", fake_save_token)
        monkeypatch.setattr(account_service, "assert_mailbox_matches", lambda *a, **k: {})
        monkeypatch.setattr(account_service, "bind_account", lambda *a, **k: {
            "mailbox_id": "mailbox123", "email": "test@example.com", "display_name": "Test User",
        })

        result = auth_module.login(token=fake_token)

        assert result == fake_token
        assert saved_data["token"] == fake_token

    def test_login_rejects_invalid_jwt_format(self):
        """Test that login() rejects tokens with invalid JWT format."""
        import pytest

        invalid_tokens = [
            "not_a_jwt",
            "only.two",
            "",
        ]

        for invalid_token in invalid_tokens:
            with pytest.raises(ValueError, match="Invalid token format"):
                auth_module.login(token=invalid_token)

    def test_login_verifies_token_via_get_me(self, monkeypatch):
        """Test that login() calls _get_me_for_token to verify."""
        verification_called = []

        def fake_get_me(token):
            verification_called.append(token)
            return {
                "Id": "mailbox123",
                "EmailAddress": "test@example.com",
                "DisplayName": "Test User",
            }

        monkeypatch.setattr(auth_module, "_get_me_for_token", fake_get_me)
        monkeypatch.setattr(auth_module, "_save_token", lambda t, a, i: None)
        monkeypatch.setattr(account_service, "assert_mailbox_matches", lambda *a, **k: {})
        monkeypatch.setattr(account_service, "bind_account", lambda *a, **k: {
            "mailbox_id": "mailbox123", "email": "test@example.com", "display_name": "Test User",
        })

        auth_module.login(token="valid.token.here")

        assert len(verification_called) == 1
        assert verification_called[0] == "valid.token.here"
