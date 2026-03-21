"""Summary command: unread inbox + today's calendar dashboard."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict
from datetime import datetime, timedelta, timezone

import click

from ._common import _get_client, _handle_api_error, _wants_json, account_option, print_summary_dashboard, to_json_envelope


def _today_window() -> tuple[str, str]:
    now_local = datetime.now().astimezone()
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    end_local = start_local + timedelta(days=1)
    return start_local.astimezone(timezone.utc).isoformat(), end_local.astimezone(timezone.utc).isoformat()


def _fetch_unread(client):
    try:
        return client.get_messages(folder="Inbox", top=5, unread_only=True)
    except Exception:
        return []


def _fetch_today_events(client):
    try:
        start, end = _today_window()
        return client.get_calendar_view(start=start, end=end, top=5)
    except Exception:
        return []


def _fetch_inbox_folder(client):
    try:
        return client.get_folder("Inbox")
    except Exception:
        return None


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def summary(as_json: bool, account_name: str | None):
    """Quick dashboard: unread inbox + today's calendar."""
    client = _get_client(account_name)

    results = {}
    with ThreadPoolExecutor(max_workers=3) as pool:
        futures = {
            pool.submit(_fetch_unread, client): "unread",
            pool.submit(_fetch_today_events, client): "events",
            pool.submit(_fetch_inbox_folder, client): "inbox",
        }
        for future in as_completed(futures):
            results[futures[future]] = future.result()

    unread_messages = results.get("unread", [])
    today_events = results.get("events", [])
    inbox_folder = results.get("inbox")

    if _wants_json(as_json):
        payload = {
            "inbox": {
                "unread_count": inbox_folder.unread_count if inbox_folder else len(unread_messages),
                "total_count": inbox_folder.total_count if inbox_folder else None,
                "messages": [asdict(message) for message in unread_messages],
            },
            "calendar": {
                "today_count": len(today_events),
                "events": [asdict(event) for event in today_events],
            },
        }
        click.echo(to_json_envelope(payload))
        return

    print_summary_dashboard(unread_messages, today_events, inbox_folder=inbox_folder)
