"""Tests for formatter output on common CLI views."""

from __future__ import annotations

from outlook_cli import formatter


def test_print_inbox_shows_categories_and_flags(capsys, make_email):
    message = make_email(
        categories=["Finance"],
        is_read=False,
        has_attachments=True,
        flag_status="flagged",
        display_num=7,
    )

    formatter.print_inbox([message])
    captured = capsys.readouterr()

    assert "Finance" in captured.err
    assert "7" in captured.err
    assert "*@!" in captured.err


def test_print_event_detail_includes_online_and_recurrence(capsys, make_event):
    event = make_event(
        is_online_meeting=True,
        online_meeting_url="https://teams.microsoft.com/l/meetup-join/123",
        recurrence={"Pattern": {"Type": "Weekly", "DaysOfWeek": ["Monday"], "Interval": 1}, "Range": {"Type": "Numbered", "NumberOfOccurrences": 4}},
        response_status="Accepted",
        event_type="SeriesMaster",
    )

    formatter.print_event_detail(event)
    captured = capsys.readouterr()

    assert "teams.microsoft.com" in captured.err
    assert "Weekly on Monday" in captured.err
    assert "Accepted" in captured.err
    assert "SeriesMaster" in captured.err


def test_print_categories_renders_category_counts(capsys):
    formatter.print_categories([{"Category": "Finance", "Color": 7, "UnreadCount": 2, "ItemCount": 10}])
    captured = capsys.readouterr()

    assert "Finance" in captured.err
    assert "2" in captured.err
    assert "10" in captured.err


def test_print_accounts_marks_current_profile(capsys):
    formatter.print_accounts(
        [
            {"name": "default", "current": True, "bound": True, "email": "a@example.com", "display_name": "Alice"},
            {"name": "work", "current": False, "bound": False, "email": None, "display_name": None},
        ]
    )
    captured = capsys.readouterr()

    assert "default" in captured.err
    assert "work" in captured.err
    assert "unbound" in captured.err
