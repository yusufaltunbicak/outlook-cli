"""Management commands: mark-read, move, copy, delete, flag."""

from __future__ import annotations

import re
from datetime import datetime, timedelta

import click

from ._common import _get_client, _handle_api_error, account_option, print_success


@click.command("mark-read")
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--unread", is_flag=True, help="Mark as unread instead")
@account_option
@_handle_api_error
def mark_read(message_ids: tuple, unread: bool, account_name: str | None):
    """Mark messages as read (or unread with --unread). Accepts multiple IDs."""
    client = _get_client()
    status = "unread" if unread else "read"
    for mid in message_ids:
        client.mark_read(mid, is_read=not unread)
        print_success(f"Message #{mid} marked as {status}")


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("destination")
@account_option
@_handle_api_error
def move(message_ids: tuple, destination: str, account_name: str | None):
    """Move messages to another folder. Accepts multiple IDs."""
    client = _get_client()
    for mid in message_ids:
        client.move_message(mid, destination)
        print_success(f"Message #{mid} moved to {destination}")


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("destination")
@account_option
@_handle_api_error
def copy(message_ids: tuple, destination: str, account_name: str | None):
    """Copy messages to another folder. Accepts multiple IDs."""
    client = _get_client()
    for mid in message_ids:
        client.copy_message(mid, destination)
        print_success(f"Message #{mid} copied to {destination}")


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def delete(message_ids: tuple, yes: bool, account_name: str | None):
    """Delete messages. Accepts multiple IDs."""
    if not yes:
        ids_str = ", ".join(f"#{m}" for m in message_ids)
        click.confirm(f"Delete {ids_str}?", abort=True)
    client = _get_client()
    for mid in message_ids:
        client.delete_message(mid)
        print_success(f"Message #{mid} deleted")


def _parse_due_date(s: str) -> str:
    """Parse a due date string into YYYY-MM-DD format.

    Supports: today, tomorrow, YYYY-MM-DD, +Nd (e.g. +3d).
    """
    s = s.strip().lower()
    today = datetime.now().date()

    if s == "today":
        return today.isoformat()
    if s == "tomorrow":
        return (today + timedelta(days=1)).isoformat()

    # +Nd relative days
    m = re.match(r'^\+(\d+)d$', s)
    if m:
        return (today + timedelta(days=int(m.group(1)))).isoformat()

    # ISO date
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return s
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: today, tomorrow, +3d, or YYYY-MM-DD"
    )


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--due", default=None, help="Due date: today, tomorrow, +3d, or YYYY-MM-DD")
@click.option("--complete", is_flag=True, help="Mark flag as complete")
@click.option("--clear", is_flag=True, help="Remove flag")
@account_option
@_handle_api_error
def flag(message_ids: tuple, due: str | None, complete: bool, clear: bool, account_name: str | None):
    """Flag messages for follow-up. Accepts multiple IDs.

    \b
    Examples:
      outlook flag 3                    # flag message
      outlook flag 3 4 5                # flag multiple messages
      outlook flag 3 --due tomorrow     # flag with due date
      outlook flag 3 --due 2026-03-20   # flag with specific date
      outlook flag 3 --due +3d          # flag due in 3 days
      outlook flag 3 --complete         # mark flag as complete
      outlook flag 3 --clear            # remove flag
    """
    if complete and clear:
        raise click.UsageError("Cannot use --complete and --clear together.")

    if complete:
        status = "complete"
    elif clear:
        status = "notFlagged"
    else:
        status = "flagged"

    due_date = _parse_due_date(due) if due else None

    client = _get_client()
    for mid in message_ids:
        client.set_flag(mid, status=status, due_date=due_date)
        if status == "flagged" and due_date:
            print_success(f"Message #{mid} flagged (due: {due_date})")
        elif status == "flagged":
            print_success(f"Message #{mid} flagged")
        elif status == "complete":
            print_success(f"Message #{mid} flag marked complete")
        else:
            print_success(f"Message #{mid} flag cleared")


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--unpin", is_flag=True, help="Unpin messages instead")
@account_option
@_handle_api_error
def pin(message_ids: tuple, unpin: bool, account_name: str | None):
    """Pin or unpin messages. Pinned messages stay at the top of your inbox.

    \b
    Examples:
      outlook pin 3              # pin message
      outlook pin 3 4 5          # pin multiple messages
      outlook pin 3 --unpin      # unpin message
    """
    client = _get_client()
    for mid in message_ids:
        client.pin_message(mid, pinned=not unpin)
        action = "unpinned" if unpin else "pinned"
        print_success(f"Message #{mid} {action}")
