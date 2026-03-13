from __future__ import annotations

from datetime import datetime, timezone

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

from .models import Attachment, Attendee, Contact, Email, Event, Folder

console = Console(stderr=True)


def print_inbox(messages: list[Email]) -> None:
    show_categories = any(msg.categories for msg in messages)

    table = Table(show_header=True, header_style="bold cyan", box=None, pad_edge=False)
    table.add_column("#", style="dim", width=5, justify="right")
    table.add_column("From", width=28, no_wrap=True)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    if show_categories:
        table.add_column("Category", width=16, no_wrap=True, overflow="ellipsis")
    table.add_column("Date", width=9, no_wrap=True, justify="right")
    table.add_column("", width=2)  # flags

    for msg in messages:
        flags = ""
        if not msg.is_read:
            flags += "*"
        if msg.has_attachments:
            flags += "@"

        style = "bold" if not msg.is_read else ""
        row = [
            str(msg.display_num),
            _truncate(str(msg.sender), 25),
            _truncate(msg.subject, 50),
        ]
        if show_categories:
            row.append(_truncate(", ".join(msg.categories), 16) if msg.categories else "")
        row.extend([_format_date(msg.received), flags])
        table.add_row(*row, style=style)

    console.print(table)


def print_email(email: Email) -> None:
    header = (
        f"[bold]From:[/bold] {email.sender}\n"
        f"[bold]To:[/bold] {', '.join(str(r) for r in email.to)}\n"
    )
    if email.cc:
        header += f"[bold]Cc:[/bold] {', '.join(str(r) for r in email.cc)}\n"
    header += f"[bold]Date:[/bold] {email.received.strftime('%Y-%m-%d %H:%M')}\n"
    if email.categories:
        header += f"[bold]Categories:[/bold] {', '.join(email.categories)}\n"
    header += f"[bold]Subject:[/bold] {email.subject}"

    body = _html_to_text(email.body) if email.body_type == "HTML" else email.body

    console.print(Panel(header, title=f"Message #{email.display_num}", border_style="cyan"))
    console.print()
    console.print(body)


def print_email_raw(email: Email) -> None:
    """Print raw HTML body."""
    console.print(email.body)


def print_folders(folders: list[Folder]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("Name", min_width=20)
    table.add_column("Unread", justify="right", width=8)
    table.add_column("Total", justify="right", width=8)

    for f in folders:
        style = "bold" if f.unread_count > 0 else ""
        table.add_row(f.name, str(f.unread_count), str(f.total_count), style=style)

    console.print(table)


def print_attachments(attachments: list[Attachment]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("#", width=4, justify="right")
    table.add_column("Name", min_width=30)
    table.add_column("Type", width=20)
    table.add_column("Size", width=10, justify="right")

    for i, att in enumerate(attachments, 1):
        table.add_row(str(i), att.name, att.content_type, _format_size(att.size))

    console.print(table)


def print_events(events: list[Event]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None, pad_edge=False)
    table.add_column("#", style="dim", width=5, justify="right")
    table.add_column("Date", width=12)
    table.add_column("Time", width=13)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    table.add_column("Location", max_width=20, no_wrap=True)
    table.add_column("Ppl", width=4, justify="right")

    for ev in events:
        if ev.is_all_day:
            time_str = "All day"
        else:
            time_str = f"{ev.start.strftime('%H:%M')}-{ev.end.strftime('%H:%M')}"
        ppl = str(len(ev.attendees)) if ev.attendees else ""
        table.add_row(
            str(ev.display_num) if ev.display_num else "",
            ev.start.strftime("%Y-%m-%d"),
            time_str,
            _truncate(ev.subject, 45),
            _truncate(ev.location, 20),
            ppl,
        )

    console.print(table)


def print_event_detail(event: Event) -> None:
    header = f"[bold]Subject:[/bold] {event.subject}\n"
    if event.is_all_day:
        header += f"[bold]When:[/bold] {event.start.strftime('%Y-%m-%d')} (All day)\n"
    else:
        header += f"[bold]Start:[/bold] {event.start.strftime('%Y-%m-%d %H:%M')}\n"
        header += f"[bold]End:[/bold] {event.end.strftime('%Y-%m-%d %H:%M')}\n"
    if event.location:
        header += f"[bold]Location:[/bold] {event.location}\n"
    header += f"[bold]Organizer:[/bold] {event.organizer}\n"
    header += f"[bold]Show as:[/bold] {event.show_as}\n"
    if event.is_online_meeting and event.online_meeting_url:
        header += f"[bold]Online:[/bold] {event.online_meeting_url}\n"
    if event.categories:
        header += f"[bold]Categories:[/bold] {', '.join(event.categories)}\n"
    if event.response_status:
        header += f"[bold]Your response:[/bold] {event.response_status}\n"
    if event.recurrence:
        header += f"[bold]Recurrence:[/bold] {_format_recurrence(event.recurrence)}\n"
    if event.event_type and event.event_type != "SingleInstance":
        header += f"[bold]Type:[/bold] {event.event_type}\n"
    if event.is_cancelled:
        header += "[bold red]CANCELLED[/bold red]\n"

    num = f"Event #{event.display_num}" if event.display_num else "Event"
    console.print(Panel(header.rstrip(), title=num, border_style="cyan"))

    if event.attendees:
        console.print(f"\n[bold]Attendees ({len(event.attendees)}):[/bold]")
        for att in event.attendees:
            resp_icon = {"Accepted": "[green]v[/green]", "Declined": "[red]x[/red]",
                         "TentativelyAccepted": "[yellow]?[/yellow]"}.get(att.response, "[dim]-[/dim]")
            att_type = f" [dim]({att.type})[/dim]" if att.type != "Required" else ""
            console.print(f"  {resp_icon} {att.email}{att_type}")

    if event.body_preview:
        console.print(f"\n{event.body_preview}")


def print_calendars(calendars: list[dict]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("Name", min_width=25)
    table.add_column("Owner", min_width=30)
    table.add_column("Color", width=12)
    table.add_column("Edit", width=5, justify="center")

    for cal in calendars:
        owner = cal.get("Owner", {}).get("Address", "")
        can_edit = "Yes" if cal.get("CanEdit") else "No"
        table.add_row(cal.get("Name", ""), owner, cal.get("Color", ""), can_edit)

    console.print(table)


def print_meeting_suggestions(suggestions: list[dict]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("#", width=4, justify="right")
    table.add_column("Start", width=18)
    table.add_column("End", width=18)
    table.add_column("Confidence", width=10, justify="right")
    table.add_column("Availability", ratio=1)

    for i, s in enumerate(suggestions, 1):
        slot = s.get("MeetingTimeSlot", {})
        start = slot.get("Start", {}).get("DateTime", "")[:16]
        end = slot.get("End", {}).get("DateTime", "")[:16]
        confidence = f"{s.get('Confidence', 0)}%"
        avail_parts = []
        for att in s.get("AttendeeAvailability", []):
            email = att.get("Attendee", {}).get("EmailAddress", {}).get("Address", "")
            avail = att.get("Availability", "?")
            avail_parts.append(f"{email}={avail}")
        table.add_row(str(i), start, end, confidence, "; ".join(avail_parts))

    console.print(table)


def print_people(people: list[dict]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("Name", min_width=25)
    table.add_column("Email", min_width=30)
    table.add_column("Title", max_width=25)

    for p in people:
        emails = p.get("ScoredEmailAddresses", [])
        email = emails[0].get("Address", "") if emails else ""
        table.add_row(
            p.get("DisplayName", ""),
            email,
            p.get("JobTitle", "") or "",
        )

    console.print(table)


def print_contacts(contacts: list[Contact]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("Name", min_width=25)
    table.add_column("Email", min_width=30)
    table.add_column("Company", max_width=20)
    table.add_column("Title", max_width=20)

    for c in contacts:
        email = c.email_addresses[0].address if c.email_addresses else ""
        table.add_row(c.display_name, email, c.company, c.job_title)

    console.print(table)


def print_categories(categories: list[dict]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("Category", min_width=25)
    table.add_column("Unread", justify="right", width=8)
    table.add_column("Total", justify="right", width=8)

    for c in categories:
        style = "bold" if c.get("UnreadCount", 0) > 0 else ""
        table.add_row(
            c["Category"],
            str(c.get("UnreadCount", 0)),
            str(c.get("ItemCount", 0)),
            style=style,
        )

    console.print(table)


def print_whoami(data: dict) -> None:
    console.print(f"[bold]Name:[/bold]  {data.get('DisplayName', 'N/A')}")
    console.print(f"[bold]Email:[/bold] {data.get('EmailAddress', 'N/A')}")
    console.print(f"[bold]Alias:[/bold] {data.get('Alias', 'N/A')}")


def print_success(msg: str) -> None:
    console.print(f"[green]{msg}[/green]")


def print_error(msg: str) -> None:
    console.print(f"[red]{msg}[/red]")


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _truncate(s: str, max_len: int) -> str:
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + "\u2026"


def _format_date(dt: datetime) -> str:
    # Compare in local timezone so "today" matches user's perception
    now_local = datetime.now().astimezone()
    if dt.tzinfo is None:
        from datetime import timezone as tz
        dt = dt.replace(tzinfo=tz.utc)
    dt_local = dt.astimezone(now_local.tzinfo)
    diff = now_local - dt_local
    if dt_local.date() == now_local.date():
        return dt_local.strftime("%H:%M")
    if (now_local.date() - dt_local.date()).days == 1:
        return "Yday"
    if diff.days < 7:
        return dt_local.strftime("%a")  # Mon, Tue...
    if dt_local.year == now_local.year:
        return dt_local.strftime("%d %b")  # 03 Mar
    return dt_local.strftime("%d %b %y")  # 03 Mar 25


def _format_size(size_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.0f}{unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f}TB"


def _format_recurrence(rec: dict) -> str:
    """Format recurrence dict into human-readable string."""
    pat = rec.get("Pattern", {})
    rng = rec.get("Range", {})
    ptype = pat.get("Type", "")
    interval = pat.get("Interval", 1)

    if ptype == "Daily":
        desc = f"Every {interval} day(s)" if interval > 1 else "Daily"
    elif ptype == "Weekly":
        days = ", ".join(pat.get("DaysOfWeek", []))
        desc = f"Every {interval} week(s) on {days}" if interval > 1 else f"Weekly on {days}"
    elif ptype == "AbsoluteMonthly":
        day = pat.get("DayOfMonth", "?")
        desc = f"Every {interval} month(s) on day {day}" if interval > 1 else f"Monthly on day {day}"
    elif ptype == "RelativeMonthly":
        idx = pat.get("Index", "")
        days = ", ".join(pat.get("DaysOfWeek", []))
        desc = f"Monthly on {idx} {days}"
    elif ptype == "AbsoluteYearly":
        month = pat.get("Month", "?")
        day = pat.get("DayOfMonth", "?")
        desc = f"Yearly on {month}/{day}"
    else:
        desc = ptype

    rtype = rng.get("Type", "")
    if rtype == "Numbered":
        desc += f" ({rng.get('NumberOfOccurrences', '?')} times)"
    elif rtype == "EndDate":
        desc += f" (until {rng.get('EndDate', '?')})"
    elif rtype == "NoEnd":
        desc += " (no end)"

    return desc


def _html_to_text(html: str) -> str:
    """Convert HTML email body to readable text."""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        # Remove style and script tags
        for tag in soup(["style", "script"]):
            tag.decompose()
        return soup.get_text(separator="\n", strip=True)
    except ImportError:
        # Fallback: strip tags with regex
        import re
        return re.sub(r"<[^>]+>", "", html)
