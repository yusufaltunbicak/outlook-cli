"""Tests for the summary dashboard command."""

from __future__ import annotations

import json

from outlook_cli.commands import summary as summary_cmd


def test_summary_outputs_combined_json(runner, tty_mode, monkeypatch, make_email, make_event, make_folder):
    fake_client = object()
    unread = [make_email(subject="Quarterly report", is_read=False, display_num=12)]
    events = [make_event(subject="Standup", display_num=30)]
    inbox = make_folder(name="Inbox", unread_count=3, total_count=20)

    monkeypatch.setattr(summary_cmd, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(summary_cmd, "_fetch_unread", lambda client: unread)
    monkeypatch.setattr(summary_cmd, "_fetch_today_events", lambda client: events)
    monkeypatch.setattr(summary_cmd, "_fetch_inbox_folder", lambda client: inbox)

    result = runner.invoke(summary_cmd.summary, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["inbox"]["unread_count"] == 3
    assert payload["data"]["inbox"]["messages"][0]["subject"] == "Quarterly report"
    assert payload["data"]["calendar"]["today_count"] == 1


def test_summary_renders_dashboard(runner, tty_mode, monkeypatch, make_email, make_event, make_folder):
    fake_client = object()
    unread = [make_email(subject="Deploy notification", is_read=False, display_num=15)]
    events = [make_event(subject="1:1 with Manager")]
    inbox = make_folder(name="Inbox", unread_count=1, total_count=10)

    monkeypatch.setattr(summary_cmd, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(summary_cmd, "_fetch_unread", lambda client: unread)
    monkeypatch.setattr(summary_cmd, "_fetch_today_events", lambda client: events)
    monkeypatch.setattr(summary_cmd, "_fetch_inbox_folder", lambda client: inbox)

    result = runner.invoke(summary_cmd.summary, [])

    assert result.exit_code == 0
    assert "1 unread" in result.output
    assert "Deploy notification" in result.output
    assert "Today's Calendar" in result.output
