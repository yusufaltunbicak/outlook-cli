#!/usr/bin/env python3
"""Generate demo SVG screenshots for README usage."""

from __future__ import annotations

import sys
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path

from rich.console import Console

ROOT = Path(__file__).resolve().parents[1]
sys.path = [str(ROOT)] + [path for path in sys.path if path != str(ROOT)]

from outlook_cli.cli import BANNER
from outlook_cli import formatter
from outlook_cli.models import Attendee, Email, EmailAddress, Event

ASSETS_DIR = ROOT / "assets"


def _email(
    *,
    display_num: int,
    sender: tuple[str, str],
    subject: str,
    received: datetime,
    categories: list[str] | None = None,
    is_read: bool = True,
    has_attachments: bool = False,
    flag_status: str = "notFlagged",
    body: str = "Thanks for reviewing this.\n\nSee the latest numbers below.",
) -> Email:
    return Email(
        id=f"msg-{display_num}",
        subject=subject,
        sender=EmailAddress(name=sender[0], address=sender[1]),
        to=[EmailAddress(name="Yusuf", address="yusuf@example.com")],
        cc=[],
        received=received,
        preview=body.splitlines()[0],
        body=body,
        body_type="Text",
        is_read=is_read,
        has_attachments=has_attachments,
        importance="Normal",
        conversation_id=f"conv-{display_num}",
        categories=categories or [],
        flag_status=flag_status,
        display_num=display_num,
    )


def _event(
    *,
    display_num: int,
    subject: str,
    start: datetime,
    end: datetime,
    location: str,
    response_status: str = "",
    is_all_day: bool = False,
) -> Event:
    return Event(
        id=f"event-{display_num}",
        subject=subject,
        start=start,
        end=end,
        location=location,
        organizer=EmailAddress(name="Alice", address="alice@example.com"),
        is_all_day=is_all_day,
        body_preview="Agenda and call details are attached.",
        body="Agenda and call details are attached.",
        body_type="Text",
        attendees=[
            Attendee(email=EmailAddress(name="Bob", address="bob@example.com"), type="Required", response="Accepted"),
        ],
        response_status=response_status,
        display_num=display_num,
    )


@contextmanager
def _recording_console(width: int = 100):
    record_console = Console(
        record=True,
        width=width,
        force_terminal=True,
        color_system="truecolor",
    )
    original_console = formatter.console
    formatter.console = record_console
    try:
        yield record_console
    finally:
        formatter.console = original_console


def _write_svg(name: str, title: str, render) -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    with _recording_console() as console:
        render(console)
        svg = console.export_svg(title=title)
    (ASSETS_DIR / f"{name}.svg").write_text(svg, encoding="utf-8")


def _render_help(console: Console) -> None:
    console.print("> outlook --help", style="bold green")
    console.print()
    console.print(f"[bold cyan]{BANNER}[/bold cyan]", highlight=False)
    console.print("  [dim]Outlook 365 from your terminal[/dim]")
    console.print()
    console.print("Usage: outlook [OPTIONS] COMMAND [ARGS]...")
    console.print()
    console.print("Commands:")
    console.print("  inbox             Show inbox messages")
    console.print("  summary           Quick dashboard: unread inbox + today's calendar")
    console.print("  read              Read an email by its display number")
    console.print("  calendar          Show calendar events")
    console.print("  search            Search messages")
    console.print("  account           Manage named Outlook account profiles")


def main() -> None:
    now = datetime.now(timezone.utc)

    unread = [
        _email(
            display_num=12,
            sender=("Alice Johnson", "alice@example.com"),
            subject="Re: Q1 Report",
            received=now - timedelta(hours=1),
            categories=["Finance"],
            is_read=False,
            flag_status="flagged",
        ),
        _email(
            display_num=15,
            sender=("Bob Smith", "bob@example.com"),
            subject="Deploy notification",
            received=now - timedelta(hours=2),
            categories=["Ops"],
            is_read=False,
            has_attachments=True,
        ),
        _email(
            display_num=18,
            sender=("HR Department", "hr@example.com"),
            subject="Benefits enrollment",
            received=now - timedelta(days=1),
            is_read=False,
        ),
    ]
    category_colors = {"Finance": 7, "Ops": 4, "VIP": 8}

    today = now.astimezone()
    base_day = today.replace(hour=0, minute=0, second=0, microsecond=0)
    events = [
        _event(
            display_num=31,
            subject="Team Standup",
            start=base_day.replace(hour=9),
            end=base_day.replace(hour=9, minute=30),
            location="Teams",
            response_status="Accepted",
        ),
        _event(
            display_num=32,
            subject="1:1 with Manager",
            start=base_day.replace(hour=14),
            end=base_day.replace(hour=15),
            location="Focus Room",
            response_status="TentativelyAccepted",
        ),
    ]

    account_rows = [
        {"name": "default", "current": True, "bound": True, "email": "yusuf@example.com", "display_name": "Yusuf Altunbicak"},
        {"name": "work", "current": False, "bound": True, "email": "yusuf@company.com", "display_name": "Yusuf (Work)"},
        {"name": "sandbox", "current": False, "bound": False, "email": None, "display_name": None},
    ]

    _write_svg("help", "outlook --help", _render_help)
    _write_svg(
        "inbox",
        "outlook inbox",
        lambda console: (
            console.print("> outlook inbox", style="bold green"),
            console.print(),
            formatter.print_inbox(unread, category_colors=category_colors),
        ),
    )
    _write_svg(
        "read",
        "outlook read 12",
        lambda console: (
            console.print("> outlook read 12", style="bold green"),
            console.print(),
            formatter.print_email(
                _email(
                    display_num=12,
                    sender=("Alice Johnson", "alice@example.com"),
                    subject="Re: Q1 Report",
                    received=now - timedelta(hours=1),
                    categories=["Finance", "VIP"],
                    body="Hi Yusuf,\n\nI've attached the revised Q1 workbook.\nPlease review before 5pm.",
                    is_read=False,
                    has_attachments=True,
                    flag_status="flagged",
                )
            ),
        ),
    )
    _write_svg(
        "calendar",
        "outlook calendar",
        lambda console: (
            console.print("> outlook calendar --days 1", style="bold green"),
            console.print(),
            formatter.print_events(events),
        ),
    )
    _write_svg(
        "summary",
        "outlook summary",
        lambda console: (
            console.print("> outlook summary", style="bold green"),
            formatter.print_summary_dashboard(unread, events),
        ),
    )
    _write_svg(
        "search",
        "outlook search report",
        lambda console: (
            console.print('> outlook search "report"', style="bold green"),
            console.print(),
            formatter.print_inbox(
                [
                    _email(
                        display_num=21,
                        sender=("Finance Bot", "finance@example.com"),
                        subject="Q1 report workbook",
                        received=now - timedelta(days=2),
                        categories=["Finance"],
                        has_attachments=True,
                    ),
                    _email(
                        display_num=22,
                        sender=("Alice Johnson", "alice@example.com"),
                        subject="Quarterly report follow-up",
                        received=now - timedelta(days=3),
                        categories=["VIP"],
                    ),
                ],
                category_colors=category_colors,
            ),
        ),
    )
    _write_svg(
        "accounts",
        "outlook account list",
        lambda console: (
            console.print("> outlook account list", style="bold green"),
            console.print(),
            formatter.print_accounts(account_rows),
        ),
    )


if __name__ == "__main__":
    main()
