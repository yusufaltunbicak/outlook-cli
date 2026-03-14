"""Schedule commands: schedule, schedule-list, schedule-cancel, schedule-draft."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import click
from rich.table import Table

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    cfg,
    console,
    print_error,
    print_success,
    to_json_envelope,
)
from .mail import _show_attachment_info


def _parse_schedule_time(s: str) -> datetime:
    """Parse schedule time from various formats.

    Supports:
      +30m, +1h, +2h30m           relative offset
      tomorrow 09:00              relative day
      today 17:00                 relative day
      2024-03-15T10:00            ISO format
      2024-03-15 10:00            space-separated
    """
    import re
    from datetime import timezone as tz

    now = datetime.now(tz.utc)

    # Relative offset: +30m, +1h, +2h30m
    offset_match = re.match(r'^\+(?:(\d+)h)?(?:(\d+)m)?$', s)
    if offset_match:
        hours = int(offset_match.group(1) or 0)
        minutes = int(offset_match.group(2) or 0)
        if hours == 0 and minutes == 0:
            raise click.BadParameter(f"Invalid offset: {s}")
        return now + timedelta(hours=hours, minutes=minutes)

    # today/tomorrow HH:MM
    day_match = re.match(r'^(today|tomorrow)\s+(\d{1,2}:\d{2})$', s, re.IGNORECASE)
    if day_match:
        day_word, time_str = day_match.groups()
        local_now = datetime.now().astimezone()
        h, m = map(int, time_str.split(":"))
        target = local_now.replace(hour=h, minute=m, second=0, microsecond=0)
        if day_word.lower() == "tomorrow":
            target += timedelta(days=1)
        return target.astimezone(tz.utc)

    # ISO-like: 2024-03-15T10:00 or 2024-03-15 10:00
    try:
        s = s.replace(" ", "T", 1) if " " in s and "T" not in s else s
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.astimezone()
        return dt.astimezone(tz.utc)
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: +30m, +1h, tomorrow 09:00, or 2024-03-15T10:00"
    )


def _print_schedule_entries(entries: list[dict]) -> None:
    table = Table(show_header=True, header_style="bold cyan", box=None, pad_edge=False)
    table.add_column("#", style="dim", width=4, justify="right")
    table.add_column("To", width=28, no_wrap=True)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    table.add_column("Scheduled", width=16, no_wrap=True, justify="right")
    table.add_column("", width=6)

    now = datetime.now(timezone.utc)
    for i, entry in enumerate(entries, 1):
        to_str = ", ".join(entry.get("to", []))
        sched = entry.get("scheduled_at", "")
        try:
            sched_dt = datetime.fromisoformat(sched.replace("Z", "+00:00"))
            local_dt = sched_dt.astimezone(datetime.now().astimezone().tzinfo)
            sched_display = local_dt.strftime("%Y-%m-%d %H:%M")
        except (ValueError, AttributeError):
            sched_display = sched
        has_draft = bool(entry.get("message_id"))
        src_tag = "[cyan]draft[/cyan]" if has_draft else "[dim]queued[/dim]"
        table.add_row(str(i), to_str[:28], entry.get("subject", "")[:50], sched_display, src_tag)

    console.print(table)


@click.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body")
@click.argument("at")
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def schedule(to: str, subject: str, body: str, at: str, cc: tuple, attach: tuple, is_html: bool, sig_name: str | None, as_json: bool, yes: bool):
    """Schedule an email to be sent later.

    AT is the scheduled send time. Accepts:
    +30m, +1h, +2h30m (relative), tomorrow 09:00, 2024-03-15T10:00
    """
    from ..signature_manager import append_signature, get_signature

    send_at = _parse_schedule_time(at)

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    to_list = [addr.strip() for addr in to.split(",")]
    cc_list = list(cc) if cc else None

    if not yes:
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        console.print(f"  [bold]To:[/bold] {', '.join(to_list)}")
        if cc_list:
            console.print(f"  [bold]CC:[/bold] {', '.join(cc_list)}")
        console.print(f"  [bold]Subject:[/bold] {subject}")
        console.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        _show_attachment_info(attach)
        console.print(f"  [bold]Scheduled:[/bold] {local_send.strftime('%Y-%m-%d %H:%M')}")
        click.confirm("Schedule this email?", abort=True)

    client = _get_client()
    send_at_str = send_at.strftime("%Y-%m-%dT%H:%M:%SZ")

    if attach:
        # Draft flow: create draft -> attach files -> schedule draft
        email = client.create_draft(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)
        client.attach_files(email.id, list(attach))
        client.schedule_draft(email.id, send_at_str)
    else:
        client.schedule_send(
            to=to_list, subject=subject, body=body, cc=cc_list,
            html=is_html, send_at=send_at_str,
        )

    if _wants_json(as_json):
        click.echo(to_json_envelope({"status": "scheduled", "to": to_list, "subject": subject, "scheduled_at": send_at.isoformat()}))
    else:
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        print_success(f"Email scheduled to {to} at {local_send.strftime('%Y-%m-%d %H:%M')}")


@click.command(name="schedule-list")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def schedule_list(as_json: bool):
    """List scheduled (pending) emails."""
    client = _get_client()
    entries = client.get_scheduled_list()

    if _wants_json(as_json):
        click.echo(to_json_envelope(entries))
    else:
        if not entries:
            print_success("No scheduled emails.")
        else:
            console.print("[bold cyan]Scheduled Emails[/bold cyan]")
            _print_schedule_entries(entries)


@click.command(name="schedule-cancel")
@click.argument("index", type=int)
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def schedule_cancel(index: int, yes: bool):
    """Cancel a scheduled email by its list number.

    For draft entries: deletes the draft from server (prevents sending).
    For queued entries: removes local tracking only.
    Run schedule-list to see numbers.
    """
    client = _get_client()
    entries = client.get_scheduled_list()
    if index < 1 or index > len(entries):
        print_error(f"Invalid index #{index}. Run 'outlook schedule-list' to see entries.")
        return

    entry = entries[index - 1]
    if not yes:
        console.print(f"  [bold]To:[/bold] {', '.join(entry['to'])}")
        console.print(f"  [bold]Subject:[/bold] {entry['subject']}")
        console.print(f"  [bold]Scheduled:[/bold] {entry['scheduled_at']}")
        click.confirm(f"Remove scheduled entry #{index}?", abort=True)

    result = client.cancel_scheduled_entry(index)
    if result and result.get("server_deleted"):
        print_success(f"Scheduled email #{index} cancelled and draft deleted: {entry['subject']}")
    else:
        print_success(f"Scheduled entry #{index} removed: {entry['subject']}")


@click.command(name="schedule-draft")
@click.argument("message_id")
@click.argument("at")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def schedule_draft(message_id: str, at: str, yes: bool):
    """Schedule an existing draft to be sent later.

    AT is the scheduled send time. Accepts:
    +30m, +1h, +2h30m (relative), tomorrow 09:00, 2024-03-15T10:00
    """
    send_at = _parse_schedule_time(at)
    client = _get_client()

    if not yes:
        email = client.get_message(message_id)
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        console.print(f"  [bold]To:[/bold] {', '.join(r.address for r in email.to)}")
        if email.cc:
            console.print(f"  [bold]CC:[/bold] {', '.join(r.address for r in email.cc)}")
        console.print(f"  [bold]Subject:[/bold] {email.subject}")
        console.print(f"  [bold]Scheduled:[/bold] {local_send.strftime('%Y-%m-%d %H:%M')}")
        click.confirm(f"Schedule draft #{message_id}?", abort=True)

    client.schedule_draft(message_id, send_at.strftime("%Y-%m-%dT%H:%M:%SZ"))

    local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
    print_success(f"Draft #{message_id} scheduled for {local_send.strftime('%Y-%m-%d %H:%M')}")
