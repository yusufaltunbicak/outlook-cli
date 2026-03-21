"""Tests for shared command helpers, config loading, and CLI registration."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock

import click
import pytest

from outlook_cli import cli as cli_module
from outlook_cli.commands import _common as common
from outlook_cli.config import DEFAULTS, _deep_merge, load_config
from outlook_cli.exceptions import AuthRequiredError, ResourceNotFoundError, TokenExpiredError


class FakeOutlookClient:
    def __init__(self, token: str, account_name: str | None = None):
        self.token = token
        self.account_name = account_name


def test_load_config_returns_defaults_when_file_missing(tmp_path):
    cfg = load_config(tmp_path / "missing.yaml")

    assert cfg["max_messages"] == DEFAULTS["max_messages"]
    assert cfg["browser"]["timeout"] == DEFAULTS["browser"]["timeout"]


def test_load_config_deep_merges_nested_values(tmp_path):
    path = tmp_path / "config.yaml"
    path.write_text(
        "max_messages: 50\n"
        "browser:\n"
        "  timeout: 300\n"
    )

    cfg = load_config(path)

    assert cfg["max_messages"] == 50
    assert cfg["browser"]["timeout"] == 300
    assert cfg["browser"]["headless"] is False


def test_deep_merge_overrides_leaf_values():
    base = {"browser": {"headless": False, "timeout": 120}, "timezone": "UTC"}
    _deep_merge(base, {"browser": {"timeout": 500}, "timezone": "Europe/Istanbul"})

    assert base == {
        "browser": {"headless": False, "timeout": 500},
        "timezone": "Europe/Istanbul",
    }


def test_get_client_caches_single_instance(monkeypatch):
    common._client_cache.clear()
    monkeypatch.setattr(common, "get_token", lambda *args, **kwargs: "abc")
    monkeypatch.setattr(common, "OutlookClient", FakeOutlookClient)
    monkeypatch.setattr(common, "_check_token_expiry", lambda token, account_name: token)

    first = common._get_client()
    second = common._get_client()

    assert first is second
    assert first.token == "abc"


def test_get_client_exits_when_auth_is_unavailable(monkeypatch):
    common._client_cache.clear()
    messages = []
    monkeypatch.setattr(common, "get_token", lambda *args, **kwargs: (_ for _ in ()).throw(AuthRequiredError("login required")))
    monkeypatch.setattr(common, "print_error", lambda msg: messages.append(msg))

    with pytest.raises(SystemExit) as exc:
        common._get_client()

    assert exc.value.code == 1
    assert messages == ["login required"]


def test_get_client_caches_per_account_profile(monkeypatch):
    common._client_cache.clear()
    monkeypatch.setattr(common, "get_token", lambda *args, **kwargs: "abc")
    monkeypatch.setattr(common, "OutlookClient", FakeOutlookClient)
    monkeypatch.setattr(common, "get_account_name", lambda account_name=None, allow_missing=False: account_name or "default")
    monkeypatch.setattr(common, "_check_token_expiry", lambda token, account_name: token)

    work = common._get_client("work")
    personal = common._get_client("personal")

    assert work is not personal
    assert work.account_name == "work"
    assert personal.account_name == "personal"


def test_get_client_refreshes_expiring_cached_client(monkeypatch):
    common._client_cache.clear()
    cached = FakeOutlookClient("old-token", account_name="default")
    cached._token = "old-token"
    common._client_cache["default"] = cached

    monkeypatch.setattr(common, "OutlookClient", FakeOutlookClient)
    monkeypatch.setattr(common.account_service, "touch_account", lambda name: None)
    monkeypatch.setattr(common, "_check_token_expiry", lambda token, account_name: "new-token")

    refreshed = common._get_client()

    assert refreshed is not cached
    assert refreshed.token == "new-token"


def test_wants_json_respects_explicit_flag(monkeypatch):
    monkeypatch.setattr(common, "_is_piped", lambda: False)
    assert common._wants_json(True) is True
    assert common._wants_json(False) is False


def test_wants_json_respects_pipe(monkeypatch):
    monkeypatch.setattr(common, "_is_piped", lambda: True)
    assert common._wants_json(False) is True


def test_handle_api_error_retries_after_relogin(monkeypatch, runner, tty_mode):
    common._client_cache = {"default": object()}
    state = {"calls": 0}
    success_messages = []
    monkeypatch.setattr(common, "do_login", lambda **kwargs: "new-token")
    monkeypatch.setattr(common, "print_success", lambda msg: success_messages.append(msg))

    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        state["calls"] += 1
        if state["calls"] == 1:
            raise TokenExpiredError("expired")
        click.echo("retried")

    result = runner.invoke(cmd, [])

    assert result.exit_code == 0
    assert state["calls"] == 2
    assert common._client_cache == {}
    assert success_messages == ["Re-login successful. Retrying..."]


def test_handle_api_error_returns_json_envelope(monkeypatch, runner, tty_mode):
    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise ResourceNotFoundError("missing message")

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "not_found"
    assert payload["error"]["message"] == "missing message"


def test_handle_api_error_reports_failed_relogin(monkeypatch, runner, tty_mode):
    monkeypatch.setattr(common, "do_login", lambda **kwargs: (_ for _ in ()).throw(RuntimeError("nope")))

    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise TokenExpiredError("expired")

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == 1
    assert '"code": "session_expired"' in result.output
    assert '"code": "auth_failed"' in result.output


def test_cli_registers_expected_commands():
    expected = {
        "login",
        "whoami",
        "account",
        "inbox",
        "read",
        "thread",
        "send",
        "draft",
        "draft-send",
        "reply",
        "reply-draft",
        "forward",
        "schedule",
        "schedule-list",
        "schedule-cancel",
        "schedule-draft",
        "search",
        "summary",
        "folders",
        "folder",
        "categories",
        "categorize",
        "uncategorize",
        "category-rename",
        "category-clear",
        "category-delete",
        "category-create",
        "signature-pull",
        "signature-list",
        "signature-show",
        "signature-delete",
        "mark-read",
        "move",
        "copy",
        "delete",
        "flag",
        "pin",
        "attachments",
        "calendar",
        "event",
        "event-create",
        "event-update",
        "event-delete",
        "event-instances",
        "event-respond",
        "calendars",
        "free-busy",
        "people-search",
        "contacts",
    }

    assert expected.issubset(set(cli_module.cli.commands))


def test_cli_help_renders_banner(runner):
    result = runner.invoke(cli_module.cli, ["--help"])

    assert result.exit_code == 0
    assert "Outlook 365 from your terminal" in result.output
    assert "summary" in result.output


def test_cli_without_args_shows_help(runner):
    result = runner.invoke(cli_module.cli, [])

    assert result.exit_code == 0
    assert "Usage:" in result.output
