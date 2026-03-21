from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timezone

from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

from .models import Attachment, Attendee, Contact, Email, Event, Folder

console = Console(stderr=True)

ACTIVE_DOT = "[green]\u25cf[/green]"
INACTIVE_DOT = "[dim]\u25cb[/dim]"

OUTLOOK_CATEGORY_COLORS = {
    0: "#f25022",
    1: "#ff8c00",
    2: "#a0522d",
    3: "#ffb900",
    4: "#107c10",
    5: "#00b7c3",
    6: "#5c8a31",
    7: "#0078d4",
    8: "#5c2d91",
    9: "#c239b3",
    10: "#e3008c",
    11: "#a4262c",
    12: "#d83b01",
    13: "#ca5010",
    14: "#986f0b",
    15: "#6b6b6b",
    16: "#498205",
    17: "#038387",
    18: "#004e8c",
    19: "#8764b8",
    20: "#881798",
    21: "#c30052",
    22: "#8e562e",
    23: "#69797e",
    24: "#485a96",
}

RESPONSE_ICONS = {
    "Accepted": ("✓", "green"),
    "TentativelyAccepted": ("?", "yellow"),
    "Declined": ("✗", "red"),
}


def print_inbox(messages: list[Email], category_colors: dict[str, int] | None = None) -> None:
    show_categories = any(msg.categories for msg in messages)

    table = _table(pad_edge=True)
    table.add_column("#", style="dim", width=5, justify="right")
    table.add_column("From", width=22, no_wrap=True)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    if show_categories:
        table.add_column("Category", width=16, overflow="ellipsis")
    table.add_column("Date", width=8, no_wrap=True, justify="right")
    table.add_column("", width=5, no_wrap=True)

    for msg in messages:
        row = [
            str(msg.display_num),
            _truncate(str(msg.sender), 25),
            _truncate(msg.subject, 50),
        ]
        if show_categories:
            row.append(_category_text(msg.categories, category_colors or {}, max_len=20))
        row.extend([_format_date(msg.received), _flag_text(msg)])
        table.add_row(*row, style="bold" if not msg.is_read else "")

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
    if email.flag_status == "flagged":
        flag_info = "Flagged"
        if email.flag_due and email.flag_due != datetime.min:
            flag_info += f" (due: {email.flag_due.strftime('%Y-%m-%d')})"
        header += f"[bold]Flag:[/bold] {flag_info}\n"
    elif email.flag_status == "complete":
        header += "[bold]Flag:[/bold] Complete\n"
    header += f"[bold]Subject:[/bold] {email.subject}"

    body = _html_to_text(email.body) if email.body_type == "HTML" else email.body

    console.print(Panel(header, title=f"Message #{email.display_num}", border_style="cyan"))
    console.print()
    console.print(body)


def print_thread(messages: list[Email]) -> None:
    console.print(f"[bold cyan]Thread ({len(messages)} messages)[/bold cyan]")
    console.print()
    for i, email in enumerate(messages):
        is_last = i == len(messages) - 1
        sender = str(email.sender)
        date = email.received.strftime("%Y-%m-%d %H:%M")
        read_marker = "" if email.is_read else " [bold cyan]*[/bold cyan]"

        header = f"[bold]#{email.display_num}[/bold]  [dim]{date}[/dim]  {sender}{read_marker}"
        console.print(header)

        body = _html_to_text(email.body) if email.body_type == "HTML" else email.body
        body = body.strip()
        if body:
            lines = body.split("\n")
            if len(lines) > 20:
                lines = lines[:20] + [f"  [dim]... ({len(lines) - 20} more lines)[/dim]"]
            for line in lines:
                console.print(f"  {line}")

        if not is_last:
            console.print(f"  [dim]{'─' * 60}[/dim]")
            console.print()


def print_email_raw(email: Email) -> None:
    console.print(email.body)


def print_folders(folders: list[Folder]) -> None:
    table = _table()
    table.add_column("Name", min_width=24)
    table.add_column("Unread", justify="right", width=8)
    table.add_column("Total", justify="right", width=8)

    for folder, depth in _ordered_folders(folders):
        prefix = "" if depth == 0 else f"{'  ' * (depth - 1)}└─ "
        style = "bold" if folder.unread_count > 0 else ""
        table.add_row(
            f"{prefix}{folder.name}",
            _unread_badge(folder.unread_count),
            str(folder.total_count),
            style=style,
        )

    console.print(table)


def print_attachments(attachments: list[Attachment]) -> None:
    table = _table()
    table.add_column("#", width=4, justify="right")
    table.add_column("Name", min_width=30)
    table.add_column("Type", width=20)
    table.add_column("Size", width=10, justify="right")

    for i, att in enumerate(attachments, 1):
        table.add_row(str(i), att.name, att.content_type, _format_size(att.size))

    console.print(table)


def print_events(events: list[Event]) -> None:
    table = _table(pad_edge=True)
    table.add_column("#", style="dim", width=5, justify="right")
    table.add_column("Date", width=12)
    table.add_column("Time", width=13)
    table.add_column("", width=2)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    table.add_column("Location", max_width=20, no_wrap=True)
    table.add_column("Ppl", width=4, justify="right")

    for ev in events:
        table.add_row(
            str(ev.display_num) if ev.display_num else "",
            ev.start.strftime("%Y-%m-%d"),
            _event_time_text(ev),
            _response_icon(ev.response_status),
            _truncate(ev.subject, 45),
            _truncate(ev.location, 20),
            str(len(ev.attendees)) if ev.attendees else "",
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
            console.print(f"  {_attendee_response_icon(att)} {att.email}{_attendee_type_suffix(att)}")

    if event.body_preview:
        console.print(f"\n{event.body_preview}")


def print_calendars(calendars: list[dict]) -> None:
    table = _table()
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
    table = _table()
    table.add_column("#", width=4, justify="right")
    table.add_column("Start", width=18)
    table.add_column("End", width=18)
    table.add_column("Confidence", width=10, justify="right")
    table.add_column("Availability", ratio=1)

    for i, suggestion in enumerate(suggestions, 1):
        slot = suggestion.get("MeetingTimeSlot", {})
        start = slot.get("Start", {}).get("DateTime", "")[:16]
        end = slot.get("End", {}).get("DateTime", "")[:16]
        confidence = f"{suggestion.get('Confidence', 0)}%"
        avail_parts = []
        for attendee in suggestion.get("AttendeeAvailability", []):
            email = attendee.get("Attendee", {}).get("EmailAddress", {}).get("Address", "")
            avail = attendee.get("Availability", "?")
            avail_parts.append(f"{email}={avail}")
        table.add_row(str(i), start, end, confidence, "; ".join(avail_parts))

    console.print(table)


def print_people(people: list[dict]) -> None:
    table = _table()
    table.add_column("Name", min_width=25)
    table.add_column("Email", min_width=30)
    table.add_column("Title", max_width=25)

    for person in people:
        emails = person.get("ScoredEmailAddresses", [])
        email = emails[0].get("Address", "") if emails else ""
        table.add_row(
            person.get("DisplayName", ""),
            email,
            person.get("JobTitle", "") or "",
        )

    console.print(table)


def print_contacts(contacts: list[Contact]) -> None:
    table = _table()
    table.add_column("Name", min_width=25)
    table.add_column("Email", min_width=30)
    table.add_column("Company", max_width=20)
    table.add_column("Title", max_width=20)

    for contact in contacts:
        email = contact.email_addresses[0].address if contact.email_addresses else ""
        table.add_row(contact.display_name, email, contact.company, contact.job_title)

    console.print(table)


def print_categories(categories: list[dict]) -> None:
    table = _table()
    table.add_column("Category", min_width=25)
    table.add_column("Unread", justify="right", width=8)
    table.add_column("Total", justify="right", width=8)

    for category in categories:
        unread_count = category.get("UnreadCount", 0)
        style = "bold" if unread_count > 0 else ""
        name = category.get("Category") or category.get("Name") or ""
        color = category.get("Color", 15)
        table.add_row(
            _category_text([name], {name: color}, max_len=25),
            _unread_badge(unread_count),
            str(category.get("ItemCount", 0)),
            style=style,
        )

    console.print(table)


def print_accounts(rows: list[dict]) -> None:
    table = _table()
    table.add_column("", width=2)
    table.add_column("Account", width=12, no_wrap=True)
    table.add_column("Email", width=22, no_wrap=True, overflow="ellipsis")
    table.add_column("Display", width=16, no_wrap=True, overflow="ellipsis")
    table.add_column("Notes", width=12, no_wrap=True)

    for row in rows:
        notes = []
        if row.get("legacy_default"):
            notes.append("legacy")
        if not row.get("bound"):
            notes.append("unbound")
        table.add_row(
            ACTIVE_DOT if row.get("current") else INACTIVE_DOT,
            row.get("name", ""),
            row.get("email") or "N/A",
            row.get("display_name") or "N/A",
            ", ".join(notes),
            style="bold" if row.get("current") else "",
        )

    console.print(table)


def print_whoami(data: dict, account_name: str | None = None) -> None:
    profile = account_name or data.get("AccountProfile")
    if profile:
        console.print(f"[bold]Account:[/bold] {ACTIVE_DOT} {profile}")
    console.print(f"[bold]Status:[/bold]  {ACTIVE_DOT} Connected")
    console.print(f"[bold]Name:[/bold]    {data.get('DisplayName', 'N/A')}")
    console.print(f"[bold]Email:[/bold]   {data.get('EmailAddress', 'N/A')}")
    console.print(f"[bold]Alias:[/bold]   {data.get('Alias', 'N/A')}")


def print_summary_dashboard(
    unread_messages: list[Email],
    today_events: list[Event],
    inbox_folder: Folder | None = None,
) -> None:
    unread_count = inbox_folder.unread_count if inbox_folder else len(unread_messages)
    event_count = len(today_events)

    console.print()
    console.print(
        f"  [bold cyan]{unread_count} unread[/bold cyan] [dim](Inbox)[/dim]"
        f"     [bold cyan]{event_count} event(s)[/bold cyan] [dim]today[/dim]"
    )

    console.print()
    console.print("  [bold]Unread[/bold]")
    if not unread_messages:
        console.print("  [dim]Inbox is clear[/dim]")
    else:
        for msg in unread_messages[:5]:
            sender = _truncate(msg.sender.name or msg.sender.address or "Unknown", 18)
            subject = _truncate(msg.subject, 28)
            console.print(
                f"  [bold cyan]*[/bold cyan] [dim]#{msg.display_num}[/dim] "
                f"{sender}  {subject}  [dim]{_format_date(msg.received)}[/dim]"
            )

    console.print()
    console.print("  [bold]Today's Calendar[/bold]")
    if not today_events:
        console.print("  [dim]No events today[/dim]")
    else:
        for event in today_events[:5]:
            console.print(f"  {_summary_event_time(event)}  {_truncate(event.subject, 42)}")
    console.print()


def print_success(msg: str) -> None:
    console.print(f"[green]{msg}[/green]")


def print_error(msg: str) -> None:
    console.print(f"[red]{msg}[/red]")


def _table(*, pad_edge: bool = True) -> Table:
    return Table(
        show_header=True,
        header_style="bold cyan",
        box=box.ROUNDED,
        border_style="dim",
        pad_edge=pad_edge,
        show_lines=False,
    )


def _category_text(categories: list[str], category_colors: dict[str, int], max_len: int) -> Text:
    text = Text()
    for index, category in enumerate(categories):
        if index:
            text.append(", ", style="dim")
        color_style = OUTLOOK_CATEGORY_COLORS.get(category_colors.get(category, 15), "dim")
        text.append("●", style=color_style)
        text.append(f" {category}")
    text.truncate(max_len, overflow="ellipsis")
    return text


def _flag_text(email: Email) -> Text:
    flags = Text()
    if not email.is_read:
        flags.append("*", style="bold cyan")
    if email.has_attachments:
        flags.append("@", style="dim")
    if email.flag_status == "flagged":
        flags.append("!", style="yellow")
    elif email.flag_status == "complete":
        flags.append("v", style="green")
    return flags


def _event_time_text(event: Event) -> Text:
    if event.is_all_day:
        return Text("All Day", style="dim")
    return Text(f"{event.start.strftime('%H:%M')}-{event.end.strftime('%H:%M')}", style="cyan")


def _response_icon(response_status: str) -> Text:
    icon, style = RESPONSE_ICONS.get(response_status, ("", ""))
    return Text(icon, style=style)


def _attendee_response_icon(attendee: Attendee) -> str:
    icon, style = RESPONSE_ICONS.get(attendee.response, ("-", "dim"))
    return f"[{style}]{icon}[/{style}]"


def _attendee_type_suffix(attendee: Attendee) -> str:
    return f" [dim]({attendee.type})[/dim]" if attendee.type != "Required" else ""


def _ordered_folders(folders: list[Folder]) -> list[tuple[Folder, int]]:
    by_id = {folder.id: folder for folder in folders}
    children: dict[str | None, list[Folder]] = defaultdict(list)
    for folder in folders:
        parent_key = folder.parent_folder_id if folder.parent_folder_id in by_id else None
        children[parent_key].append(folder)

    ordered: list[tuple[Folder, int]] = []
    visited: set[str] = set()

    def walk(parent_id: str | None, depth: int) -> None:
        for child in sorted(children.get(parent_id, []), key=lambda item: item.name.lower()):
            if child.id in visited:
                continue
            visited.add(child.id)
            ordered.append((child, depth))
            walk(child.id, depth + 1)

    walk(None, 0)
    for folder in folders:
        if folder.id not in visited:
            ordered.append((folder, 0))
    return ordered


def _unread_badge(count: int) -> Text:
    if count > 1:
        return Text(f" {count} ", style="bold white on blue")
    if count == 1:
        return Text("*", style="bold cyan")
    return Text("0", style="dim")


def _summary_event_time(event: Event) -> str:
    if event.is_all_day:
        return "[dim]All Day[/dim]"
    return f"[cyan]{event.start.strftime('%H:%M')}-{event.end.strftime('%H:%M')}[/cyan]"


def _truncate(s: str, max_len: int) -> str:
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + "\u2026"


def _format_date(dt: datetime) -> str:
    now_local = datetime.now().astimezone()
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    dt_local = dt.astimezone(now_local.tzinfo)
    diff = now_local - dt_local
    if dt_local.date() == now_local.date():
        return dt_local.strftime("%H:%M")
    if (now_local.date() - dt_local.date()).days == 1:
        return "Yday"
    if diff.days < 7:
        return dt_local.strftime("%a")
    if dt_local.year == now_local.year:
        return dt_local.strftime("%d %b")
    return dt_local.strftime("%d %b %y")


def _format_size(size_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.0f}{unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f}TB"


def _format_recurrence(rec: dict) -> str:
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
    try:
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html, "html.parser")
        for tag in soup(["style", "script"]):
            tag.decompose()
        return soup.get_text(separator="\n", strip=True)
    except ImportError:
        import re

        return re.sub(r"<[^>]+>", "", html)
