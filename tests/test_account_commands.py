"""Tests for account profile CLI commands."""

from __future__ import annotations

import json

from outlook_cli.commands import account as account_cmd


def test_account_add_runs_login_without_switching_current(runner, tty_mode, monkeypatch):
    calls = []
    messages = []
    monkeypatch.setattr(account_cmd.account_service, "load_registry", lambda: {"accounts": {}, "current_account": "default"})
    monkeypatch.setattr(account_cmd, "do_login", lambda **kwargs: calls.append(kwargs) or "token")
    monkeypatch.setattr(account_cmd, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(account_cmd.account, ["add", "work"])

    assert result.exit_code == 0
    assert calls == [{"account_name": "work", "allow_create": True}]
    assert messages == ["Account profile 'work' added."]


def test_account_list_outputs_json(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(
        account_cmd.account_service,
        "list_accounts",
        lambda: [{"name": "default", "current": True, "bound": True, "email": "a@example.com", "display_name": "Alice", "legacy_default": True}],
    )

    result = runner.invoke(account_cmd.account, ["list", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["name"] == "default"
    assert payload["data"][0]["current"] is True


def test_account_switch_updates_current_profile(runner, tty_mode, monkeypatch):
    calls = []
    messages = []
    monkeypatch.setattr(account_cmd.account_service, "set_current_account", lambda name: calls.append(name))
    monkeypatch.setattr(account_cmd, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(account_cmd.account, ["switch", "work"])

    assert result.exit_code == 0
    assert calls == ["work"]
    assert messages == ["Switched to account 'work'."]


def test_account_current_outputs_json(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(account_cmd.account_service, "get_current_account_name", lambda: "work")
    monkeypatch.setattr(
        account_cmd.account_service,
        "list_accounts",
        lambda: [{"name": "work", "current": True, "bound": True, "email": "work@example.com", "display_name": "Work", "legacy_default": False}],
    )

    result = runner.invoke(account_cmd.account, ["current", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["name"] == "work"
    assert payload["data"]["email"] == "work@example.com"


def test_account_remove_deletes_profile(runner, tty_mode, monkeypatch):
    calls = []
    messages = []
    monkeypatch.setattr(account_cmd.account_service, "remove_account", lambda name: calls.append(name))
    monkeypatch.setattr(account_cmd, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(account_cmd.account, ["remove", "work", "-y"])

    assert result.exit_code == 0
    assert calls == ["work"]
    assert messages == ["Account profile 'work' removed."]
