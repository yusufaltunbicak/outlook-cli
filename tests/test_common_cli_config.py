"""Tests for shared command helpers, config loading, and CLI registration."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock

import click
import pytest

from outlook_cli import cli as cli_module
from outlook_cli.commands import _common as common
from outlook_cli.commands import manage, schedule
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
    monkeypatch.setattr(common, "_is_piped", lambda: False)
    monkeypatch.setattr(common, "get_token", lambda *args, **kwargs: (_ for _ in ()).throw(AuthRequiredError("login required")))
    monkeypatch.setattr(common, "print_error", lambda msg: messages.append(msg))

    with pytest.raises(click.exceptions.Exit) as exc:
        common._get_client()

    assert exc.value.exit_code == 4
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

    assert result.exit_code == 5
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

    assert result.exit_code == 4
    payload = json.loads(result.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "auth_failed"
    assert "Token expired. Attempting re-login..." in result.stderr


def test_handle_api_error_json_retry_keeps_stdout_single_payload(monkeypatch, runner, tty_mode):
    common._client_cache = {"default": object()}
    state = {"calls": 0}
    monkeypatch.setattr(common, "do_login", lambda **kwargs: "new-token")

    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        state["calls"] += 1
        if state["calls"] == 1:
            raise TokenExpiredError("expired")
        click.echo(common.to_json_envelope({"status": "ok"}))

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == 0
    assert state["calls"] == 2
    payload = json.loads(result.stdout)
    assert payload["ok"] is True
    assert payload["data"]["status"] == "ok"
    assert result.stderr.count("Token expired. Attempting re-login...") == 1
    assert result.stderr.count("Re-login successful. Retrying...") == 1


def test_handle_api_error_pipe_mode_retry_keeps_stdout_single_payload(monkeypatch, runner):
    common._client_cache = {"default": object()}
    state = {"calls": 0}
    monkeypatch.setattr(common, "_is_piped", lambda: True)
    monkeypatch.setattr(common, "do_login", lambda **kwargs: "new-token")

    @click.command()
    @common._handle_api_error
    def cmd():
        state["calls"] += 1
        if state["calls"] == 1:
            raise TokenExpiredError("expired")
        click.echo(common.to_json_envelope({"status": "ok"}))

    result = runner.invoke(cmd, [])

    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["ok"] is True
    assert payload["data"]["status"] == "ok"
    assert "Token expired. Attempting re-login..." in result.stderr


def test_check_token_expiry_json_mode_writes_status_to_stderr(monkeypatch, capsys):
    monkeypatch.setattr(common, "_is_json_mode", lambda: True)
    monkeypatch.setattr(common, "_decode_exp", lambda token: 0)
    monkeypatch.delenv("OUTLOOK_TOKEN", raising=False)
    monkeypatch.setattr(common, "do_login", lambda **kwargs: "fresh-token")

    token = common._check_token_expiry("stale-token", "default")

    captured = capsys.readouterr()
    assert token == "fresh-token"
    assert captured.out == ""
    assert "Token expiring soon. Re-authenticating..." in captured.err


def test_confirm_action_rejects_non_interactive_without_yes(monkeypatch):
    monkeypatch.setattr(common, "_stdin_is_tty", lambda: False)

    with pytest.raises(click.UsageError, match="Refusing to delete #1 without --yes"):
        common.confirm_action("Delete #1?", action="delete #1")


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
        "open",
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


def test_cli_no_input_after_subcommand_rejects_delete(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(manage, "_get_client", lambda: fake_client)

    result = runner.invoke(cli_module.cli, ["delete", "1", "--no-input"])

    assert result.exit_code == 2
    assert "Refusing to delete #1 without --yes" in result.output
    fake_client.delete_message.assert_not_called()


def test_cli_no_input_after_subcommand_rejects_schedule(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    result = runner.invoke(
        cli_module.cli,
        ["schedule", "a@example.com", "Subject", "Body", "+30m", "--no-input"],
    )

    assert result.exit_code == 2
    assert "Refusing to schedule this email without --yes" in result.output
    fake_client.schedule_send.assert_not_called()


def test_cli_no_input_allows_yes_bypass(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(manage, "_get_client", lambda: fake_client)

    result = runner.invoke(cli_module.cli, ["--no-input", "delete", "1", "-y"])

    assert result.exit_code == 0
    fake_client.delete_message.assert_called_once_with("1")


def test_cli_dry_run_after_subcommand_outputs_json_and_skips_send(runner, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(cli_module.mail_mod, "_get_client", lambda: fake_client)
    monkeypatch.setitem(cli_module.mail_mod.cfg, "default_signature", None)

    result = runner.invoke(
        cli_module.cli,
        ["send", "a@example.com", "Subject", "Body", "--dry-run", "--json"],
    )

    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["ok"] is True
    assert payload["data"]["dry_run"] is True
    assert payload["data"]["op"] == "send"
    assert payload["data"]["request"]["to"] == ["a@example.com"]
    fake_client.send_mail.assert_not_called()


def test_cli_dry_run_human_output_skips_delete(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(manage, "_get_client", lambda: fake_client)

    result = runner.invoke(cli_module.cli, ["delete", "1", "--dry-run"])

    assert result.exit_code == 0
    assert "Dry run: would delete" in result.output
    fake_client.delete_message.assert_not_called()


def test_cli_dry_run_bypasses_no_input_confirmation(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(manage, "_get_client", lambda: fake_client)

    result = runner.invoke(cli_module.cli, ["delete", "1", "--dry-run", "--no-input"])

    assert result.exit_code == 0
    assert "Dry run: would delete" in result.output
    fake_client.delete_message.assert_not_called()


def test_cli_enable_commands_allows_specific_command(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_me": lambda self: {"DisplayName": "Alice"}})()
    monkeypatch.setattr(cli_module.auth_mod, "_get_client", lambda: fake_client)
    monkeypatch.setattr(cli_module.auth_mod, "get_account_name", lambda account_name=None: "default")

    result = runner.invoke(cli_module.cli, ["--enable-commands", "whoami", "whoami", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["DisplayName"] == "Alice"


def test_cli_enable_commands_rewrites_option_after_subcommand(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_me": lambda self: {"DisplayName": "Alice"}})()
    monkeypatch.setattr(cli_module.auth_mod, "_get_client", lambda: fake_client)
    monkeypatch.setattr(cli_module.auth_mod, "get_account_name", lambda account_name=None: "default")

    result = runner.invoke(cli_module.cli, ["whoami", "--enable-commands", "whoami", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["DisplayName"] == "Alice"


def test_cli_enable_commands_blocks_other_commands(runner, tty_mode):
    result = runner.invoke(cli_module.cli, ["--enable-commands", "whoami", "inbox"])

    assert result.exit_code == 2
    assert "Command 'inbox' is not enabled." in result.output


def test_cli_enable_commands_all_keyword_allows_commands(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_me": lambda self: {"DisplayName": "Alice"}})()
    monkeypatch.setattr(cli_module.auth_mod, "_get_client", lambda: fake_client)
    monkeypatch.setattr(cli_module.auth_mod, "get_account_name", lambda account_name=None: "default")

    result = runner.invoke(cli_module.cli, ["--enable-commands", "all", "whoami", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["DisplayName"] == "Alice"
