"""CLI integration tests for calendar commands and recurrence helpers."""

from __future__ import annotations

import json

import pytest

from outlook_cli.commands import calendar as calendar_cmd


def test_parse_event_time_accepts_iso_like_strings():
    assert calendar_cmd._parse_event_time("2026-03-15 10:00") == "2026-03-15T10:00:00"


def test_parse_event_time_rejects_invalid_values():
    with pytest.raises(Exception):
        calendar_cmd._parse_event_time("not-a-date")


def test_build_recurrence_supports_weekly_and_monthly():
    weekly = calendar_cmd._build_recurrence(
        "weekly",
        "2026-03-15T10:00:00",
        interval=2,
        count=5,
        days="Monday,Wednesday",
    )
    monthly = calendar_cmd._build_recurrence("monthly", "2026-03-15T10:00:00", interval=1, count=3)

    assert weekly["Pattern"]["DaysOfWeek"] == ["Monday", "Wednesday"]
    assert weekly["Range"]["NumberOfOccurrences"] == 5
    assert monthly["Pattern"]["DayOfMonth"] == 15


def test_calendar_command_outputs_json(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_calendar_view = lambda **_kwargs: [make_event(subject="Standup")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--days", "3", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Standup"


def test_calendar_command_converts_table_times_to_requested_timezone(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_calendar_view = lambda **_kwargs: [make_event()]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--days", "1", "--timezone", "UTC+2"])

    assert result.exit_code == 0
    assert "12:00-13:00" in result.output


def test_event_command_converts_table_times_to_requested_timezone(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda _event_id: make_event()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event, ["1", "--timezone", "UTC+2"])

    assert result.exit_code == 0
    assert "2026-03-17 12:00" in result.output


def test_event_create_builds_recurrence_and_calls_client(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    seen = {}

    def create_event(**kwargs):
        seen.update(kwargs)
        return make_event(subject=kwargs["subject"], recurrence=kwargs["recurrence"])

    fake_client.create_event = create_event
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "Europe/Istanbul")

    result = runner.invoke(
        calendar_cmd.event_create,
        [
            "Planning",
            "2026-03-20T10:00",
            "2026-03-20T11:00",
            "--attendee",
            "a@example.com",
            "--repeat",
            "weekly",
            "--repeat-days",
            "Monday,Wednesday",
            "--repeat-count",
            "4",
            "-y",
        ],
    )

    assert result.exit_code == 0
    assert seen["timezone"] == "Europe/Istanbul"
    assert seen["attendees"] == ["a@example.com"]
    assert seen["recurrence"]["Pattern"]["Type"] == "Weekly"
    assert seen["recurrence"]["Range"]["NumberOfOccurrences"] == 4


def test_event_update_can_modify_fields_and_attendees(runner, tty_mode, monkeypatch, make_event):
    class FakeClient:
        def __init__(self):
            self.added = []
            self.removed = []
            self.updated = None

        def add_event_attendees(self, event_id, attendees):
            self.added.append((event_id, attendees))

        def remove_event_attendees(self, event_id, attendees):
            self.removed.append((event_id, attendees))

        def update_event(self, event_id, **kwargs):
            self.updated = (event_id, kwargs)
            return make_event(subject="Updated")

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_update,
        [
            "3",
            "--subject",
            "Updated",
            "--start",
            "2026-03-21T09:00",
            "--body",
            "Notes",
            "--add-attendee",
            "a@example.com",
            "--remove-attendee",
            "b@example.com",
        ],
    )

    assert result.exit_code == 0
    assert fake_client.added == [("3", ["a@example.com"])]
    assert fake_client.removed == [("3", ["b@example.com"])]
    assert fake_client.updated[0] == "3"
    assert fake_client.updated[1]["subject"] == "Updated"
    assert fake_client.updated[1]["body"] == "Notes"


def test_event_delete_series_uses_series_master_id(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda event_id: make_event(id="occurrence-id", event_type="Occurrence", series_master_id="series-id")
    fake_client._delete = lambda path: setattr(fake_client, "deleted_path", path)
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_delete, ["3", "--series", "-y"])

    assert result.exit_code == 0
    assert fake_client.deleted_path == "/events/series-id"


def test_event_instances_outputs_json(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event_instances = lambda *args, **kwargs: [make_event(subject="Occurrence")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_instances, ["3", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Occurrence"


def test_event_respond_maps_tentative_to_tentativelyaccept(runner, tty_mode, monkeypatch):
    class FakeClient:
        def respond_to_event(self, event_id, response, comment="", send_response=True):
            self.called = (event_id, response, comment, send_response)

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_respond, ["5", "tentative", "--comment", "Maybe", "--silent"])

    assert result.exit_code == 0
    assert fake_client.called == ("5", "tentativelyaccept", "Maybe", False)


def test_free_busy_uses_client_and_returns_json(runner, tty_mode, monkeypatch):
    class FakeClient:
        def find_meeting_times(self, **kwargs):
            self.kwargs = kwargs
            return [{"Confidence": 90}]

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(calendar_cmd.free_busy, ["a@example.com,b@example.com", "2026-03-20", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["Confidence"] == 90
    assert fake_client.kwargs["attendees"] == ["a@example.com", "b@example.com"]


def test_people_search_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.search_people = lambda query, top=10: [{"DisplayName": "Alice"}]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.people_search, ["alice", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["DisplayName"] == "Alice"
