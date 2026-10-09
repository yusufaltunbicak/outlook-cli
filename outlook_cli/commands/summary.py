"""Dashboard with explicit failed sections and bounded-display versus total counts."""
from __future__ import annotations
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone

import click

from ._common import _get_client, _handle_api_error, _wants_json, account_option, print_summary_dashboard, to_json_envelope
from ..exceptions import error_code_for_exception


def _today_window() -> tuple[str, str]:
    start = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
    return start.astimezone(timezone.utc).isoformat(), (start + timedelta(days=1)).astimezone(timezone.utc).isoformat()


def _fetch_unread(client):
    return client.get_messages(folder="Inbox", top=5, unread_only=True)


def _fetch_today_events(client):
    start, end = _today_window()
    return client.get_calendar_view(start=start, end=end, top=5)


def _fetch_inbox_folder(client):
    return client.get_folder("Inbox")


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def summary(as_json: bool, account_name: str | None):
    """Show unread inbox and today's events, with partial failures reported."""
    client = _get_client(account_name)
    results, errors = {}, {}
    with ThreadPoolExecutor(max_workers=3) as pool:
        futures = {pool.submit(fn, client): section for fn, section in ((_fetch_unread, "unread"), (_fetch_today_events, "events"), (_fetch_inbox_folder, "inbox"))}
        for future in as_completed(futures):
            section = futures[future]
            try:
                results[section] = future.result()
            except Exception as exc:
                errors[section] = {"code": error_code_for_exception(exc), "message": str(exc)}
    unread = results.get("unread")
    events = results.get("events")
    inbox = results.get("inbox")
    event_meta = getattr(events, "meta", {})
    # Without exhausted pagination or a server count the total is unknown.
    total_events = len(events) if events is not None and event_meta.get("complete") is True else None
    payload = {
        "inbox": {"unread_count": inbox.unread_count if inbox is not None else None, "total_count": inbox.total_count if inbox is not None else None, "displayed_count": len(unread) if unread is not None else None, "messages": unread},
        "calendar": {"today_count": total_events, "total_count": total_events, "displayed_count": len(events) if events is not None else None, "events": events, "meta": event_meta},
        "errors": errors,
    }
    if _wants_json(as_json):
        click.echo(to_json_envelope(payload, ok=not errors, meta={"partial": bool(errors and results), "failed_sections": sorted(errors)}, error={"code": "partial_failure" if results else "fetch_failed", "message": "One or more dashboard sections could not be loaded"} if errors else None))
    else:
        if results:
            print_summary_dashboard(unread or [], events or [], inbox_folder=inbox)
        for section, error in errors.items():
            click.echo(f"{section}: unavailable ({error['message']})", err=True)
    if errors:
        raise click.exceptions.Exit(1)
