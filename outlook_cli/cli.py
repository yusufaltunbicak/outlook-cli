from __future__ import annotations

import base64
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import click

from .auth import get_token, login as do_login, verify_token
from .client import OutlookClient, TokenExpiredError
from .config import load_config
from .formatter import (
    console,
    print_attachments,
    print_calendars,
    print_categories,
    print_contacts,
    print_email,
    print_email_raw,
    print_error,
    print_event_detail,
    print_events,
    print_folders,
    print_inbox,
    print_meeting_suggestions,
    print_people,
    print_success,
    print_whoami,
)
from .serialization import save_json, to_json

cfg = load_config()

# Cache client instance per session so auto-relogin can invalidate it
_client_cache: dict[str, OutlookClient] = {}


def _get_client() -> OutlookClient:
    if "c" not in _client_cache:
        try:
            token = get_token()
        except RuntimeError as e:
            print_error(str(e))
            sys.exit(1)
        _client_cache["c"] = OutlookClient(token)
    return _client_cache["c"]


def _handle_api_error(fn):
    """Decorator to catch common API errors. Auto re-login on 401."""
    import functools

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except TokenExpiredError:
            print_error("Token expired. Attempting re-login...")
            try:
                token = do_login()
                print_success("Re-login successful. Retrying...")
                # Invalidate cached client so next _get_client() uses new token
                _client_cache.clear()
                return fn(*args, **kwargs)
            except Exception:
                print_error("Auto re-login failed. Run: outlook login --force")
                sys.exit(1)
        except ValueError as e:
            print_error(str(e))
            sys.exit(1)
        except Exception as e:
            print_error(f"Error: {e}")
            sys.exit(1)

    return wrapper


@click.group()
@click.version_option(package_name="outlook-cli")
def cli():
    """Outlook 365 CLI - read, send, and manage emails from the terminal."""
    pass


# ------------------------------------------------------------------
# Auth
# ------------------------------------------------------------------

@cli.command()
@click.option("--force", is_flag=True, help="Force re-login, ignore saved session")
@click.option("--debug", is_flag=True, help="Show debug info about captured requests")
def login(force: bool, debug: bool):
    """Authenticate via browser and cache the token."""
    try:
        token = do_login(force=force, debug=debug)
        if verify_token(token):
            print_success("Logged in successfully. Token cached.")
        else:
            print_error("Login completed but token verification failed.")
    except RuntimeError as e:
        print_error(str(e))
        sys.exit(1)


@cli.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def whoami(as_json: bool):
    """Show current user info."""
    client = _get_client()
    data = client.get_me()
    if as_json:
        click.echo(to_json(data))
    else:
        print_whoami(data)


# ------------------------------------------------------------------
# Mail - Read
# ------------------------------------------------------------------

@cli.command()
@click.option("--max", "-n", "max_count", default=None, type=int, help="Number of messages")
@click.option("--unread", is_flag=True, help="Show only unread messages")
@click.option("--from", "from_filter", default=None, help="Filter by sender (name or email)")
@click.option("--subject", default=None, help="Filter by subject")
@click.option("--after", default=None, help="After date (YYYY-MM-DD)")
@click.option("--before", default=None, help="Before date (YYYY-MM-DD)")
@click.option("--has-attachments", is_flag=True, help="Only messages with attachments")
@click.option("--category", default=None, help="Filter by category name")
@click.option("--no-category", "no_category", is_flag=True, help="Only uncategorized messages")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@_handle_api_error
def inbox(
    max_count: int | None,
    unread: bool,
    from_filter: str | None,
    subject: str | None,
    after: str | None,
    before: str | None,
    has_attachments: bool,
    category: str | None,
    no_category: bool,
    as_json: bool,
    output: str | None,
):
    """Show inbox messages."""
    client = _get_client()
    top = max_count or cfg["max_messages"]
    has_filters = any([unread, from_filter, subject, after, before, has_attachments, category, no_category])

    messages = client.get_messages(
        folder="Inbox",
        top=top,
        unread_only=unread,
        filter_from=from_filter,
        filter_subject=subject,
        filter_after=after,
        filter_before=before,
        filter_has_attachments=has_attachments,
        filter_category=category,
        filter_no_category=no_category,
    )

    if as_json:
        text = to_json(messages)
        if output:
            save_json(messages, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(text)
    else:
        # Show folder summary header
        if not has_filters:
            try:
                folder_info = client.get_folder("Inbox")
                console.print(
                    f"[bold cyan]Inbox[/bold cyan]  "
                    f"[dim]{folder_info.unread_count} unread / {folder_info.total_count} total[/dim]"
                )
            except Exception:
                pass
        if not messages:
            print_success("No messages found.")
        else:
            print_inbox(messages)


@cli.command()
@click.argument("message_id")
@click.option("--raw", is_flag=True, help="Show raw HTML body")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def read(message_id: str, raw: bool, as_json: bool):
    """Read an email by its display number."""
    client = _get_client()
    email = client.get_message(message_id)

    if as_json:
        click.echo(to_json(email))
    elif raw:
        print_email_raw(email)
    else:
        print_email(email)

    # Auto mark as read
    if not email.is_read:
        try:
            client.mark_read(message_id)
        except Exception:
            pass


# ------------------------------------------------------------------
# Mail - Write
# ------------------------------------------------------------------

@cli.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body")
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@_handle_api_error
def send(to: str, subject: str, body: str, cc: tuple, is_html: bool, sig_name: str | None, as_json: bool, yes: bool):
    """Send an email. TO can be comma-separated for multiple recipients."""
    from .signature_manager import append_signature, get_signature

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    to_list = [addr.strip() for addr in to.split(",")]
    cc_list = list(cc) if cc else None

    if not yes:
        console.print(f"  [bold]To:[/bold] {', '.join(to_list)}")
        if cc_list:
            console.print(f"  [bold]CC:[/bold] {', '.join(cc_list)}")
        console.print(f"  [bold]Subject:[/bold] {subject}")
        console.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        click.confirm("Send this email?", abort=True)

    client = _get_client()
    client.send_mail(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)

    if as_json:
        click.echo(to_json({"status": "sent", "to": to_list, "subject": subject}))
    else:
        print_success(f"Mail sent to {to}")


@cli.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body")
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def draft(to: str, subject: str, body: str, cc: tuple, is_html: bool, sig_name: str | None, as_json: bool):
    """Create a draft email without sending. TO can be comma-separated."""
    from .signature_manager import append_signature, get_signature

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    client = _get_client()
    to_list = [addr.strip() for addr in to.split(",")]
    cc_list = list(cc) if cc else None
    email = client.create_draft(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)

    if as_json:
        click.echo(to_json(email))
    else:
        print_success(f"Draft created: {subject} (to: {to})")


@cli.command(name="draft-send")
@click.argument("message_id")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@_handle_api_error
def draft_send(message_id: str, yes: bool):
    """Send an existing draft by its message number."""
    client = _get_client()
    if not yes:
        email = client.get_message(message_id)
        console.print(f"  [bold]To:[/bold] {', '.join(r.address for r in email.to)}")
        if email.cc:
            console.print(f"  [bold]CC:[/bold] {', '.join(r.address for r in email.cc)}")
        console.print(f"  [bold]Subject:[/bold] {email.subject}")
        click.confirm(f"Send draft #{message_id}?", abort=True)
    client.send_draft(message_id)
    print_success(f"Draft #{message_id} sent")


@cli.command()
@click.argument("message_id")
@click.argument("body")
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@_handle_api_error
def reply(message_id: str, body: str, reply_all: bool, yes: bool):
    """Reply to an email."""
    client = _get_client()
    if not yes:
        action = "Reply all" if reply_all else "Reply"
        console.print(f"  [bold]{action} to #{message_id}[/bold]")
        console.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        click.confirm("Send this reply?", abort=True)
    client.reply(message_id, body, reply_all=reply_all)
    action = "Reply all" if reply_all else "Reply"
    print_success(f"{action} sent for message #{message_id}")


@cli.command(name="reply-draft")
@click.argument("message_id")
@click.argument("body", default="")
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--html", "is_html", is_flag=True, help="Body is HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def reply_draft(message_id: str, body: str, reply_all: bool, is_html: bool, sig_name: str | None, as_json: bool):
    """Create a reply draft without sending."""
    from .signature_manager import append_signature, get_signature

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name and body:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    client = _get_client()
    email = client.create_reply_draft(message_id, comment=body, reply_all=reply_all, html=is_html)
    action = "Reply-all" if reply_all else "Reply"
    if as_json:
        click.echo(to_json(email))
    else:
        print_success(f"{action} draft created for message #{message_id}")


@cli.command()
@click.argument("message_id")
@click.argument("to")
@click.option("--comment", "-c", default="", help="Add a comment to the forwarded message")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@_handle_api_error
def forward(message_id: str, to: str, comment: str, yes: bool):
    """Forward an email."""
    to_list = [addr.strip() for addr in to.split(",")]
    if not yes:
        console.print(f"  [bold]Forward #{message_id} to:[/bold] {', '.join(to_list)}")
        if comment:
            console.print(f"  [bold]Comment:[/bold] {comment[:100]}{'...' if len(comment) > 100 else ''}")
        click.confirm("Forward this email?", abort=True)
    client = _get_client()
    client.forward(message_id, to_list, comment=comment)
    print_success(f"Message #{message_id} forwarded to {to}")


# ------------------------------------------------------------------
# Scheduled send
# ------------------------------------------------------------------

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
            # Assume local timezone
            dt = dt.astimezone()
        return dt.astimezone(tz.utc)
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: +30m, +1h, tomorrow 09:00, or 2024-03-15T10:00"
    )


def _print_schedule_entries(entries: list[dict]) -> None:
    from rich.table import Table
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


@cli.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body")
@click.argument("at")
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def schedule(to: str, subject: str, body: str, at: str, cc: tuple, is_html: bool, sig_name: str | None, as_json: bool, yes: bool):
    """Schedule an email to be sent later.

    AT is the scheduled send time. Accepts:
    +30m, +1h, +2h30m (relative), tomorrow 09:00, 2024-03-15T10:00
    """
    from .signature_manager import append_signature, get_signature

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
        console.print(f"  [bold]Scheduled:[/bold] {local_send.strftime('%Y-%m-%d %H:%M')}")
        click.confirm("Schedule this email?", abort=True)

    client = _get_client()
    client.schedule_send(
        to=to_list, subject=subject, body=body, cc=cc_list,
        html=is_html, send_at=send_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
    )

    if as_json:
        click.echo(to_json({"status": "scheduled", "to": to_list, "subject": subject, "scheduled_at": send_at.isoformat()}))
    else:
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        print_success(f"Email scheduled to {to} at {local_send.strftime('%Y-%m-%d %H:%M')}")


@cli.command(name="schedule-list")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def schedule_list(as_json: bool):
    """List scheduled (pending) emails."""
    client = _get_client()
    entries = client.get_scheduled_list()

    if as_json:
        click.echo(to_json(entries))
    else:
        if not entries:
            print_success("No scheduled emails.")
        else:
            console.print("[bold cyan]Scheduled Emails[/bold cyan]")
            _print_schedule_entries(entries)


@cli.command(name="schedule-cancel")
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


@cli.command(name="schedule-draft")
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


# ------------------------------------------------------------------
# Search
# ------------------------------------------------------------------

@cli.command()
@click.argument("query")
@click.option("--max", "-n", "max_count", default=25, type=int, help="Max results")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@_handle_api_error
def search(query: str, max_count: int, as_json: bool, output: str | None):
    """Search messages."""
    client = _get_client()
    messages = client.search_messages(query, top=max_count)

    if as_json:
        if output:
            save_json(messages, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json(messages))
    else:
        if not messages:
            print_error("No results found.")
        else:
            print_inbox(messages)


# ------------------------------------------------------------------
# Folders
# ------------------------------------------------------------------

@cli.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@_handle_api_error
def folders(as_json: bool, output: str | None):
    """List mail folders."""
    client = _get_client()
    folder_list = client.get_folders()

    if as_json:
        if output:
            save_json(folder_list, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json(folder_list))
    else:
        print_folders(folder_list)


@cli.command()
@click.argument("name")
@click.option("--max", "-n", "max_count", default=None, type=int, help="Number of messages")
@click.option("--unread", is_flag=True, help="Show only unread messages")
@click.option("--from", "from_filter", default=None, help="Filter by sender")
@click.option("--subject", default=None, help="Filter by subject")
@click.option("--after", default=None, help="After date (YYYY-MM-DD)")
@click.option("--before", default=None, help="Before date (YYYY-MM-DD)")
@click.option("--has-attachments", is_flag=True, help="Only messages with attachments")
@click.option("--category", default=None, help="Filter by category name")
@click.option("--no-category", "no_category", is_flag=True, help="Only uncategorized messages")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def folder(
    name: str,
    max_count: int | None,
    unread: bool,
    from_filter: str | None,
    subject: str | None,
    after: str | None,
    before: str | None,
    has_attachments: bool,
    category: str | None,
    no_category: bool,
    as_json: bool,
):
    """Show messages in a specific folder."""
    client = _get_client()
    top = max_count or cfg["max_messages"]
    messages = client.get_messages(
        folder=name,
        top=top,
        unread_only=unread,
        filter_from=from_filter,
        filter_subject=subject,
        filter_after=after,
        filter_before=before,
        filter_has_attachments=has_attachments,
        filter_category=category,
        filter_no_category=no_category,
    )

    if as_json:
        click.echo(to_json(messages))
    else:
        if not messages:
            print_success(f"No messages found in '{name}'.")
        else:
            console.print(f"[bold cyan]Folder: {name}[/bold cyan]")
            print_inbox(messages)


# ------------------------------------------------------------------
# Categories
# ------------------------------------------------------------------

@cli.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def categories(as_json: bool):
    """List master categories with unread/total counts."""
    client = _get_client()
    resp = client.get_master_categories()
    cat_list = resp.get("Body", {}).get("CategoryDetailsList", [])

    if as_json:
        click.echo(to_json(cat_list))
    else:
        if not cat_list:
            print_success("No categories defined.")
        else:
            print_categories(cat_list)


@cli.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("category")
@_handle_api_error
def categorize(message_ids: tuple, category: str):
    """Add a category to messages. Accepts multiple IDs."""
    client = _get_client()
    for mid in message_ids:
        result = client.add_category(mid, category)
        print_success(f"Message #{mid} categorized as: {', '.join(result)}")


@cli.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("category")
@_handle_api_error
def uncategorize(message_ids: tuple, category: str):
    """Remove a category from messages. Accepts multiple IDs."""
    client = _get_client()
    for mid in message_ids:
        result = client.remove_category(mid, category)
        if result:
            print_success(f"Message #{mid} categories: {', '.join(result)}")
        else:
            print_success(f"Message #{mid} has no categories.")


@cli.command("category-rename")
@click.argument("old_name")
@click.argument("new_name")
@click.option("--no-propagate", is_flag=True, help="Only rename master category, skip updating messages")
@_handle_api_error
def category_rename(old_name: str, new_name: str, no_propagate: bool):
    """Rename a master category and update all messages."""
    from .category_manager import rename_category

    def on_progress(done, _total):
        console.print(f"  [dim]{done} messages updated...[/dim]")

    token = get_token()
    count = rename_category(token, old_name, new_name, propagate=not no_propagate, on_progress=on_progress)
    print_success(f"Renamed '{old_name}' → '{new_name}'")
    if count:
        print_success(f"  {count} messages updated")


@cli.command("category-clear")
@click.argument("name")
@click.option("--folder", default=None, help="Limit to a specific folder")
@click.option("--max", "-n", "max_messages", type=int, default=None, help="Max messages to clear")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def category_clear(name: str, folder: str | None, max_messages: int | None, yes: bool):
    """Remove a category label from messages (does not delete master category)."""
    from .category_manager import clear_category

    scope = f"in '{folder}'" if folder else "in all folders"
    limit = f" (max {max_messages})" if max_messages else ""
    if not yes:
        click.confirm(f"Remove '{name}' from messages {scope}{limit}?", abort=True)

    def on_progress(done, _total):
        console.print(f"  [dim]{done} messages cleared...[/dim]")

    token = get_token()
    count = clear_category(token, name, folder=folder, max_messages=max_messages, on_progress=on_progress)
    print_success(f"Cleared '{name}' from {count} messages")


@cli.command("category-delete")
@click.argument("name")
@click.option("--no-propagate", is_flag=True, help="Only delete master category, skip clearing messages")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def category_delete(name: str, no_propagate: bool, yes: bool):
    """Delete a master category and remove it from all messages."""
    from .category_manager import clear_category, delete_category

    if not yes:
        click.confirm(f"Delete category '{name}' and remove from all messages?", abort=True)

    token = get_token()

    if not no_propagate:
        def on_progress(done, _total):
            console.print(f"  [dim]{done} messages cleared...[/dim]")

        count = clear_category(token, name, on_progress=on_progress)
        if count:
            print_success(f"  Cleared from {count} messages")

    delete_category(token, name)
    print_success(f"Deleted category '{name}'")


@cli.command("category-create")
@click.argument("name")
@click.option("--color", type=int, default=15, help="Color index (0-24)")
@_handle_api_error
def category_create(name: str, color: int):
    """Create a new master category."""
    from .category_manager import create_category
    token = get_token()
    create_category(token, name, color=color)
    print_success(f"Created category '{name}'")


# ------------------------------------------------------------------
# Signatures
# ------------------------------------------------------------------

@cli.command("signature-pull")
@click.option("--name", "-n", default=None, help="Name for the signature (default: auto-detect)")
@_handle_api_error
def signature_pull(name: str | None):
    """Extract your signature from a recent sent email and save it."""
    from .signature_manager import pull_signature, save_signature

    token = get_token()
    sig_html, source_subject = pull_signature(token)

    if not name:
        name = click.prompt("Signature name", default="default")

    path = save_signature(name, sig_html)
    print_success(f"Signature '{name}' saved from: {source_subject}")
    console.print(f"  [dim]{path}[/dim]")


@cli.command("signature-list")
def signature_list():
    """List saved signatures."""
    from .signature_manager import list_signatures

    sigs = list_signatures()
    if not sigs:
        print_success("No signatures saved. Run 'outlook signature-pull' to extract one.")
    else:
        for s in sigs:
            default = " [bold cyan](default)[/bold cyan]" if s == cfg.get("default_signature") else ""
            console.print(f"  {s}{default}")


@cli.command("signature-show")
@click.argument("name")
@_handle_api_error
def signature_show(name: str):
    """Preview a saved signature."""
    from .signature_manager import get_signature

    from bs4 import BeautifulSoup

    sig_html = get_signature(name)
    text = BeautifulSoup(sig_html, "html.parser").get_text("\n", strip=True)
    console.print(text)


@cli.command("signature-delete")
@click.argument("name")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def signature_delete(name: str, yes: bool):
    """Delete a saved signature."""
    from .signature_manager import delete_signature

    if not yes:
        click.confirm(f"Delete signature '{name}'?", abort=True)
    delete_signature(name)
    print_success(f"Deleted signature '{name}'")


# ------------------------------------------------------------------
# Management
# ------------------------------------------------------------------

@cli.command("mark-read")
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--unread", is_flag=True, help="Mark as unread instead")
@_handle_api_error
def mark_read(message_ids: tuple, unread: bool):
    """Mark messages as read (or unread with --unread). Accepts multiple IDs."""
    client = _get_client()
    status = "unread" if unread else "read"
    for mid in message_ids:
        client.mark_read(mid, is_read=not unread)
        print_success(f"Message #{mid} marked as {status}")


@cli.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("destination")
@_handle_api_error
def move(message_ids: tuple, destination: str):
    """Move messages to another folder. Accepts multiple IDs."""
    client = _get_client()
    for mid in message_ids:
        client.move_message(mid, destination)
        print_success(f"Message #{mid} moved to {destination}")


@cli.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def delete(message_ids: tuple, yes: bool):
    """Delete messages. Accepts multiple IDs."""
    if not yes:
        ids_str = ", ".join(f"#{m}" for m in message_ids)
        click.confirm(f"Delete {ids_str}?", abort=True)
    client = _get_client()
    for mid in message_ids:
        client.delete_message(mid)
        print_success(f"Message #{mid} deleted")


# ------------------------------------------------------------------
# Attachments
# ------------------------------------------------------------------

@cli.command()
@click.argument("message_id")
@click.option("-d", "--download", is_flag=True, help="Download all attachments")
@click.option("--save-to", type=click.Path(), default=".", help="Download directory")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def attachments(message_id: str, download: bool, save_to: str, as_json: bool):
    """List or download attachments for a message."""
    client = _get_client()
    atts = client.get_attachments(message_id)

    if not atts:
        print_success("No attachments.")
        return

    if as_json:
        click.echo(to_json(atts))
        return

    print_attachments(atts)

    if download:
        save_path = Path(save_to)
        save_path.mkdir(parents=True, exist_ok=True)
        for att in atts:
            if att.content_bytes:
                file_path = save_path / att.name
                file_path.write_bytes(base64.b64decode(att.content_bytes))
                print_success(f"  Saved: {file_path}")
            else:
                # Need to fetch full attachment with content
                full = client.download_attachment(message_id, att.id)
                if full.content_bytes:
                    file_path = save_path / full.name
                    file_path.write_bytes(base64.b64decode(full.content_bytes))
                    print_success(f"  Saved: {file_path}")


# ------------------------------------------------------------------
# Calendar
# ------------------------------------------------------------------

@cli.command()
@click.option("--days", default=7, type=int, help="Number of days to show")
@click.option("--calendar", "cal_name", default=None, help="Calendar name (default: your primary calendar)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@_handle_api_error
def calendar(days: int, cal_name: str | None, as_json: bool, output: str | None):
    """Show upcoming calendar events."""
    client = _get_client()
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days)
    events = client.get_calendar_view(
        start=now.isoformat(),
        end=end.isoformat(),
        calendar_name=cal_name,
    )

    if as_json:
        if output:
            save_json(events, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json(events))
    else:
        if not events:
            print_success(f"No events in the next {days} days.")
        else:
            console.print(f"[bold cyan]Calendar ({days} days)[/bold cyan]")
            print_events(events)


@cli.command()
@click.argument("event_id")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def event(event_id: str, as_json: bool):
    """View event details by display number."""
    client = _get_client()
    ev = client.get_event(event_id)
    if as_json:
        click.echo(to_json(ev))
    else:
        print_event_detail(ev)


def _parse_event_time(s: str) -> str:
    """Parse event time to ISO format for the API.

    Supports:
      +2h, +30m, +1h30m           relative offset
      tomorrow 09:00              relative day
      today 17:00                 relative day
      2026-03-15T10:00            ISO format
      2026-03-15 10:00            space-separated
    """
    import re
    now = datetime.now()

    # Relative offset: +30m, +1h, +2h30m
    offset_match = re.match(r'^\+(?:(\d+)h)?(?:(\d+)m)?$', s)
    if offset_match:
        hours = int(offset_match.group(1) or 0)
        minutes = int(offset_match.group(2) or 0)
        if hours == 0 and minutes == 0:
            raise click.BadParameter(f"Invalid offset: {s}")
        target = now + timedelta(hours=hours, minutes=minutes)
        return target.strftime("%Y-%m-%dT%H:%M:%S")

    # today/tomorrow HH:MM
    day_match = re.match(r'^(today|tomorrow)\s+(\d{1,2}:\d{2})$', s, re.IGNORECASE)
    if day_match:
        day_word, time_str = day_match.groups()
        h, m = map(int, time_str.split(":"))
        target = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if day_word.lower() == "tomorrow":
            target += timedelta(days=1)
        return target.strftime("%Y-%m-%dT%H:%M:%S")

    # ISO-like: 2026-03-15T10:00 or 2026-03-15 10:00
    s_norm = s.replace(" ", "T", 1) if " " in s and "T" not in s else s
    try:
        dt = datetime.fromisoformat(s_norm)
        return dt.strftime("%Y-%m-%dT%H:%M:%S")
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: +1h, tomorrow 09:00, or 2026-03-15T10:00"
    )


def _build_recurrence(
    repeat: str,
    start_dt: str,
    interval: int = 1,
    count: int | None = None,
    until: str | None = None,
    days: str | None = None,
) -> dict:
    """Build Outlook API Recurrence payload."""
    start_date = start_dt[:10]  # YYYY-MM-DD
    day_of_week = datetime.fromisoformat(start_dt).strftime("%A")

    # Pattern
    if repeat == "daily":
        pattern = {"Type": "Daily", "Interval": interval}
    elif repeat == "weekly":
        if days:
            day_list = [d.strip() for d in days.split(",")]
        else:
            day_list = [day_of_week]
        pattern = {"Type": "Weekly", "Interval": interval, "DaysOfWeek": day_list}
    elif repeat == "monthly":
        day_of_month = int(start_dt[8:10])
        pattern = {"Type": "AbsoluteMonthly", "Interval": interval, "DayOfMonth": day_of_month}
    else:
        raise click.BadParameter(f"Unknown repeat type: {repeat}")

    # Range
    if count:
        rng = {"Type": "Numbered", "StartDate": start_date, "NumberOfOccurrences": count}
    elif until:
        rng = {"Type": "EndDate", "StartDate": start_date, "EndDate": until}
    else:
        # Default: 4 occurrences if nothing specified
        rng = {"Type": "Numbered", "StartDate": start_date, "NumberOfOccurrences": 4}

    return {"Pattern": pattern, "Range": rng}


@cli.command("event-create")
@click.argument("subject")
@click.argument("start")
@click.argument("end")
@click.option("--attendee", "-a", multiple=True, help="Attendee email (repeatable)")
@click.option("--location", "-l", default=None, help="Event location")
@click.option("--body", "-b", default=None, help="Event body/description")
@click.option("--html", "is_html", is_flag=True, help="Body is HTML")
@click.option("--all-day", is_flag=True, help="All-day event")
@click.option("--reminder", type=int, default=15, help="Reminder minutes before (default 15)")
@click.option("--teams", is_flag=True, help="Create as Teams online meeting")
@click.option("--repeat", type=click.Choice(["daily", "weekly", "monthly"]), default=None, help="Recurrence pattern")
@click.option("--repeat-interval", type=int, default=1, help="Repeat every N days/weeks/months (default 1)")
@click.option("--repeat-count", type=int, default=None, help="Number of occurrences")
@click.option("--repeat-until", default=None, help="End date for recurrence (YYYY-MM-DD)")
@click.option("--repeat-days", default=None, help="Days of week for weekly (comma-separated: Monday,Wednesday,Friday)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def event_create(
    subject: str, start: str, end: str,
    attendee: tuple, location: str | None, body: str | None,
    is_html: bool, all_day: bool, reminder: int, teams: bool,
    repeat: str | None, repeat_interval: int, repeat_count: int | None,
    repeat_until: str | None, repeat_days: str | None,
    as_json: bool, yes: bool,
):
    """Create a calendar event.

    START and END accept: +1h, tomorrow 09:00, 2026-03-15T10:00

    Recurrence examples:
      --repeat weekly --repeat-count 4
      --repeat daily --repeat-until 2026-04-01
      --repeat weekly --repeat-days Monday,Wednesday --repeat-count 8
      --repeat monthly --repeat-count 3
    """
    start_dt = _parse_event_time(start)
    end_dt = _parse_event_time(end)
    attendees = list(attendee) if attendee else None

    # Build recurrence payload
    recurrence = None
    if repeat:
        recurrence = _build_recurrence(
            repeat, start_dt, interval=repeat_interval,
            count=repeat_count, until=repeat_until, days=repeat_days,
        )

    if not yes:
        console.print(f"  [bold]Subject:[/bold] {subject}")
        console.print(f"  [bold]Start:[/bold] {start_dt}")
        console.print(f"  [bold]End:[/bold] {end_dt}")
        if attendees:
            console.print(f"  [bold]Attendees:[/bold] {', '.join(attendees)}")
        if location:
            console.print(f"  [bold]Location:[/bold] {location}")
        if recurrence:
            pat = recurrence["Pattern"]
            rng = recurrence["Range"]
            console.print(f"  [bold]Repeat:[/bold] {pat['Type']} every {pat['Interval']}")
            if rng["Type"] == "Numbered":
                console.print(f"  [bold]Occurrences:[/bold] {rng['NumberOfOccurrences']}")
            elif rng["Type"] == "EndDate":
                console.print(f"  [bold]Until:[/bold] {rng['EndDate']}")
        click.confirm("Create this event?", abort=True)

    client = _get_client()
    ev = client.create_event(
        subject=subject, start=start_dt, end=end_dt,
        timezone=cfg.get("timezone", "UTC"),
        attendees=attendees, location=location,
        body=body, html=is_html, is_all_day=all_day,
        reminder_minutes=reminder, is_online_meeting=teams,
        recurrence=recurrence,
    )

    if as_json:
        click.echo(to_json(ev))
    else:
        print_success(f"Event created: {ev.subject}")
        console.print(f"  [dim]{ev.start.strftime('%Y-%m-%d %H:%M')} - {ev.end.strftime('%H:%M')}[/dim]")
        if ev.attendees:
            console.print(f"  [dim]Attendees: {len(ev.attendees)}[/dim]")
        if ev.recurrence:
            from .formatter import _format_recurrence
            console.print(f"  [dim]Recurrence: {_format_recurrence(ev.recurrence)}[/dim]")


@cli.command("event-update")
@click.argument("event_id")
@click.option("--subject", "-s", default=None, help="New subject")
@click.option("--start", default=None, help="New start time")
@click.option("--end", default=None, help="New end time")
@click.option("--location", "-l", default=None, help="New location")
@click.option("--body", "-b", default=None, help="New body/description")
@click.option("--add-attendee", multiple=True, help="Add attendee email (repeatable)")
@click.option("--remove-attendee", multiple=True, help="Remove attendee email (repeatable)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def event_update(
    event_id: str, subject: str | None, start: str | None, end: str | None,
    location: str | None, body: str | None,
    add_attendee: tuple, remove_attendee: tuple, as_json: bool,
):
    """Update a calendar event."""
    client = _get_client()

    # Handle attendee add/remove separately
    if add_attendee:
        client.add_event_attendees(event_id, list(add_attendee))
        print_success(f"Added {len(add_attendee)} attendee(s) to event #{event_id}")
    if remove_attendee:
        client.remove_event_attendees(event_id, list(remove_attendee))
        print_success(f"Removed {len(remove_attendee)} attendee(s) from event #{event_id}")

    # Handle field updates
    kwargs: dict = {}
    if subject:
        kwargs["subject"] = subject
    if start:
        kwargs["start"] = _parse_event_time(start)
    if end:
        kwargs["end"] = _parse_event_time(end)
    if location:
        kwargs["location"] = location
    if body:
        kwargs["body"] = body

    if kwargs:
        kwargs["timezone"] = cfg.get("timezone", "UTC")
        ev = client.update_event(event_id, **kwargs)
        if as_json:
            click.echo(to_json(ev))
        else:
            print_success(f"Event #{event_id} updated: {ev.subject}")
    elif not add_attendee and not remove_attendee:
        print_error("No changes specified. Use --subject, --start, --end, --location, --body, --add-attendee, --remove-attendee.")


@cli.command("event-delete")
@click.argument("event_ids", nargs=-1, required=True)
@click.option("--series", is_flag=True, help="Delete entire recurring series (uses SeriesMasterId)")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def event_delete(event_ids: tuple, series: bool, yes: bool):
    """Delete calendar events. Accepts multiple IDs.

    For recurring events: deletes single occurrence by default.
    Use --series to delete the entire series.
    """
    client = _get_client()
    for eid in event_ids:
        if series:
            # Get the event to find its series master ID
            ev = client.get_event(eid)
            if ev.series_master_id:
                target_id = ev.series_master_id
                label = f"entire series of #{eid}"
            elif ev.event_type == "SeriesMaster":
                target_id = ev.id
                label = f"series #{eid}"
            else:
                target_id = ev.id
                label = f"event #{eid} (not a recurring event)"
            if not yes:
                click.confirm(f"Delete {label}?", abort=True)
            # Delete using raw ID since series master may not be in id_map
            client._delete(f"/events/{target_id}")
            print_success(f"Deleted {label}")
        else:
            if not yes:
                click.confirm(f"Delete event #{eid}?", abort=True)
            client.delete_event(eid)
            print_success(f"Event #{eid} deleted")


@cli.command("event-instances")
@click.argument("event_id")
@click.option("--days", default=90, type=int, help="Look-ahead days (default 90)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def event_instances(event_id: str, days: int, as_json: bool):
    """List occurrences of a recurring event."""
    client = _get_client()
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days)
    events = client.get_event_instances(
        event_id,
        start=now.isoformat(),
        end=end.isoformat(),
    )
    if as_json:
        click.echo(to_json(events))
    else:
        if not events:
            print_success("No occurrences found.")
        else:
            console.print(f"[bold cyan]Occurrences ({len(events)})[/bold cyan]")
            print_events(events)


@cli.command("event-respond")
@click.argument("event_id")
@click.argument("response", type=click.Choice(["accept", "decline", "tentative"]))
@click.option("--comment", "-c", default="", help="Response comment")
@click.option("--silent", is_flag=True, help="Don't send response to organizer")
@_handle_api_error
def event_respond(event_id: str, response: str, comment: str, silent: bool):
    """Respond to a meeting invitation (accept/decline/tentative)."""
    response_map = {
        "accept": "accept",
        "decline": "decline",
        "tentative": "tentativelyaccept",
    }
    client = _get_client()
    client.respond_to_event(
        event_id, response_map[response],
        comment=comment, send_response=not silent,
    )
    print_success(f"Event #{event_id}: {response}")


@cli.command(name="calendars")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def calendars_cmd(as_json: bool):
    """List available calendars."""
    client = _get_client()
    cals = client.get_calendars()
    if as_json:
        click.echo(to_json(cals))
    else:
        if not cals:
            print_success("No calendars found.")
        else:
            print_calendars(cals)


@cli.command("free-busy")
@click.argument("attendees")
@click.argument("date")
@click.option("--start-hour", default=9, type=int, help="Start hour (default 9)")
@click.option("--end-hour", default=18, type=int, help="End hour (default 18)")
@click.option("--duration", "-d", default=60, type=int, help="Meeting duration in minutes (default 60)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def free_busy(attendees: str, date: str, start_hour: int, end_hour: int, duration: int, as_json: bool):
    """Find available meeting times.

    ATTENDEES: comma-separated emails. DATE: YYYY-MM-DD, today, or tomorrow.
    """
    addr_list = [a.strip() for a in attendees.split(",")]

    # Parse date
    if date.lower() == "today":
        d = datetime.now()
    elif date.lower() == "tomorrow":
        d = datetime.now() + timedelta(days=1)
    else:
        d = datetime.fromisoformat(date)

    start_str = d.replace(hour=start_hour, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S")
    end_str = d.replace(hour=end_hour, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S")

    client = _get_client()
    suggestions = client.find_meeting_times(
        attendees=addr_list, start=start_str, end=end_str,
        duration_minutes=duration,
        timezone=cfg.get("timezone", "UTC"),
    )

    if as_json:
        click.echo(to_json(suggestions))
    else:
        if not suggestions:
            print_error("No available meeting slots found.")
        else:
            console.print(f"[bold cyan]Available slots ({len(suggestions)})[/bold cyan]")
            print_meeting_suggestions(suggestions)


@cli.command("people-search")
@click.argument("query")
@click.option("--max", "-n", "max_count", default=10, type=int, help="Max results")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def people_search(query: str, max_count: int, as_json: bool):
    """Search people by name for attendee autocomplete."""
    client = _get_client()
    results = client.search_people(query, top=max_count)
    if as_json:
        click.echo(to_json(results))
    else:
        if not results:
            print_error("No people found.")
        else:
            print_people(results)


# ------------------------------------------------------------------
# Contacts
# ------------------------------------------------------------------

@cli.command()
@click.option("--max", "-n", "max_count", default=50, type=int, help="Max contacts")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@_handle_api_error
def contacts(max_count: int, as_json: bool, output: str | None):
    """List contacts."""
    client = _get_client()
    contact_list = client.get_contacts(top=max_count)

    if as_json:
        if output:
            save_json(contact_list, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json(contact_list))
    else:
        if not contact_list:
            print_success("No contacts found.")
        else:
            print_contacts(contact_list)
