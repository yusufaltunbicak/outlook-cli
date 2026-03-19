"""Calendar commands: calendar, event, event-create/update/delete/instances/respond, calendars, free-busy, people-search."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    console,
    print_calendars,
    print_error,
    print_event_detail,
    print_events,
    print_meeting_suggestions,
    print_people,
    print_success,
    save_json,
    to_json_envelope,
)


def _parse_timezone(tz_str: str | None):
    """Parse timezone string to timezone object.

    Supports:
      None          -> None (no conversion)
      UTC           -> UTC
      UTC+8, UTC-5  -> fixed offset
      Asia/Shanghai -> IANA timezone name
    """
    if tz_str is None:
        return None

    tz_str = tz_str.strip()

    if tz_str.upper() == "UTC":
        return timezone.utc

    import re
    offset_match = re.match(r'^UTC([+-])(\d{1,2})(?::(\d{2}))?$', tz_str, re.IGNORECASE)
    if offset_match:
        sign = 1 if offset_match.group(1) == '+' else -1
        hours = int(offset_match.group(2))
        minutes = int(offset_match.group(3) or 0)
        return timezone(sign * timedelta(hours=hours, minutes=minutes))

    try:
        import zoneinfo
        return zoneinfo.ZoneInfo(tz_str)
    except (ImportError, AttributeError):
        try:
            from dateutil import tz
            return tz.gettz(tz_str)
        except ImportError:
            raise click.BadParameter(
                f"Unknown timezone: {tz_str}. Install python-dateutil for IANA timezone support."
            )
    except Exception:
        raise click.BadParameter(f"Unknown timezone: {tz_str}")


def _resolve_output_tz(tz_str: str | None):
    """Resolve output timezone from --timezone flag or config.yaml."""
    if tz_str is not None:
        return _parse_timezone(tz_str)
    config_tz = cfg.get("timezone", "UTC")
    if config_tz and config_tz.upper() != "UTC":
        return _parse_timezone(config_tz)
    return None


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

    if count:
        rng = {"Type": "Numbered", "StartDate": start_date, "NumberOfOccurrences": count}
    elif until:
        rng = {"Type": "EndDate", "StartDate": start_date, "EndDate": until}
    else:
        rng = {"Type": "Numbered", "StartDate": start_date, "NumberOfOccurrences": 4}

    return {"Pattern": pattern, "Range": rng}


@click.command()
@click.option("--days", default=7, type=int, help="Number of days to show (negative for past)")
@click.option("--calendar", "cal_name", default=None, help="Calendar name (default: your primary calendar)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def calendar(days: int, cal_name: str | None, as_json: bool, tz_str: str | None, output: str | None, account_name: str | None):
    """Show calendar events (past or future).

    Datetimes are automatically converted to your system's local timezone.
    Use --timezone to override with a specific timezone.

    Examples:
      outlook calendar --days 7      # Next 7 days
      outlook calendar --days -7     # Past 7 days
      outlook calendar --days -30    # Past 30 days

    Note: Date ranges are calculated in local timezone, not UTC.
    This means --days -1 will show events from yesterday 00:00 to today 00:00 (local time).
    """
    tz = _resolve_output_tz(tz_str)

    import datetime as dt
    now_local = dt.datetime.now().astimezone()
    today_midnight = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

    # Calculate date range in local timezone
    if days >= 0:
        start = today_midnight
        end = today_midnight + timedelta(days=days)
        range_desc = f"next {days} days"
    else:
        start = today_midnight + timedelta(days=days)  # days is negative
        end = today_midnight
        range_desc = f"past {-days} days"

    # Convert to UTC for API
    start_utc = start.astimezone(timezone.utc)
    end_utc = end.astimezone(timezone.utc)

    client = _get_client()
    events = client.get_calendar_view(
        start=start_utc.isoformat(),
        end=end_utc.isoformat(),
        calendar_name=cal_name,
    )

    if _wants_json(as_json):
        if output:
            save_json(events, output, tz=tz)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(events, tz=tz))
    else:
        if not events:
            print_success(f"No events in the {range_desc}.")
        else:
            console.print(f"[bold cyan]Calendar ({range_desc})[/bold cyan]")
            print_events(events)


@click.command()
@click.argument("event_id")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@account_option
@_handle_api_error
def event(event_id: str, as_json: bool, tz_str: str | None, account_name: str | None):
    """View event details by display number."""
    tz = _resolve_output_tz(tz_str)
    client = _get_client()
    ev = client.get_event(event_id)
    if _wants_json(as_json):
        click.echo(to_json_envelope(ev, tz=tz))
    else:
        print_event_detail(ev)


@click.command("event-create")
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
@account_option
@_handle_api_error
def event_create(
    subject: str, start: str, end: str,
    attendee: tuple, location: str | None, body: str | None,
    is_html: bool, all_day: bool, reminder: int, teams: bool,
    repeat: str | None, repeat_interval: int, repeat_count: int | None,
    repeat_until: str | None, repeat_days: str | None,
    as_json: bool, yes: bool, account_name: str | None,
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

    if _wants_json(as_json):
        click.echo(to_json_envelope(ev))
    else:
        print_success(f"Event created: {ev.subject}")
        console.print(f"  [dim]{ev.start.strftime('%Y-%m-%d %H:%M')} - {ev.end.strftime('%H:%M')}[/dim]")
        if ev.attendees:
            console.print(f"  [dim]Attendees: {len(ev.attendees)}[/dim]")
        if ev.recurrence:
            from ..formatter import _format_recurrence
            console.print(f"  [dim]Recurrence: {_format_recurrence(ev.recurrence)}[/dim]")


@click.command("event-update")
@click.argument("event_id")
@click.option("--subject", "-s", default=None, help="New subject")
@click.option("--start", default=None, help="New start time")
@click.option("--end", default=None, help="New end time")
@click.option("--location", "-l", default=None, help="New location")
@click.option("--body", "-b", default=None, help="New body/description")
@click.option("--add-attendee", multiple=True, help="Add attendee email (repeatable)")
@click.option("--remove-attendee", multiple=True, help="Remove attendee email (repeatable)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def event_update(
    event_id: str, subject: str | None, start: str | None, end: str | None,
    location: str | None, body: str | None,
    add_attendee: tuple, remove_attendee: tuple, as_json: bool, account_name: str | None,
):
    """Update a calendar event."""
    client = _get_client()

    if add_attendee:
        client.add_event_attendees(event_id, list(add_attendee))
        print_success(f"Added {len(add_attendee)} attendee(s) to event #{event_id}")
    if remove_attendee:
        client.remove_event_attendees(event_id, list(remove_attendee))
        print_success(f"Removed {len(remove_attendee)} attendee(s) from event #{event_id}")

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
        if _wants_json(as_json):
            click.echo(to_json_envelope(ev))
        else:
            print_success(f"Event #{event_id} updated: {ev.subject}")
    elif not add_attendee and not remove_attendee:
        print_error("No changes specified. Use --subject, --start, --end, --location, --body, --add-attendee, --remove-attendee.")


@click.command("event-delete")
@click.argument("event_ids", nargs=-1, required=True)
@click.option("--series", is_flag=True, help="Delete entire recurring series (uses SeriesMasterId)")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def event_delete(event_ids: tuple, series: bool, yes: bool, account_name: str | None):
    """Delete calendar events. Accepts multiple IDs.

    For recurring events: deletes single occurrence by default.
    Use --series to delete the entire series.
    """
    client = _get_client()
    for eid in event_ids:
        if series:
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
            client._delete(f"/events/{target_id}")
            print_success(f"Deleted {label}")
        else:
            if not yes:
                click.confirm(f"Delete event #{eid}?", abort=True)
            client.delete_event(eid)
            print_success(f"Event #{eid} deleted")


@click.command("event-instances")
@click.argument("event_id")
@click.option("--days", default=90, type=int, help="Look-ahead days (default 90)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@account_option
@_handle_api_error
def event_instances(event_id: str, days: int, as_json: bool, tz_str: str | None, account_name: str | None):
    """List occurrences of a recurring event."""
    tz = _resolve_output_tz(tz_str)
    client = _get_client()
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days)
    events = client.get_event_instances(
        event_id,
        start=now.isoformat(),
        end=end.isoformat(),
    )
    if _wants_json(as_json):
        click.echo(to_json_envelope(events, tz=tz))
    else:
        if not events:
            print_success("No occurrences found.")
        else:
            console.print(f"[bold cyan]Occurrences ({len(events)})[/bold cyan]")
            print_events(events)


@click.command("event-respond")
@click.argument("event_id")
@click.argument("response", type=click.Choice(["accept", "decline", "tentative"]))
@click.option("--comment", "-c", default="", help="Response comment")
@click.option("--silent", is_flag=True, help="Don't send response to organizer")
@account_option
@_handle_api_error
def event_respond(event_id: str, response: str, comment: str, silent: bool, account_name: str | None):
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


@click.command(name="calendars")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def calendars_cmd(as_json: bool, account_name: str | None):
    """List available calendars."""
    client = _get_client()
    cals = client.get_calendars()
    if _wants_json(as_json):
        click.echo(to_json_envelope(cals))
    else:
        if not cals:
            print_success("No calendars found.")
        else:
            print_calendars(cals)


@click.command("free-busy")
@click.argument("attendees")
@click.argument("date")
@click.option("--start-hour", default=9, type=int, help="Start hour (default 9)")
@click.option("--end-hour", default=18, type=int, help="End hour (default 18)")
@click.option("--duration", "-d", default=60, type=int, help="Meeting duration in minutes (default 60)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@account_option
@_handle_api_error
def free_busy(attendees: str, date: str, start_hour: int, end_hour: int, duration: int, as_json: bool, tz_str: str | None, account_name: str | None):
    """Find available meeting times.

    ATTENDEES: comma-separated emails. DATE: YYYY-MM-DD, today, or tomorrow.
    """
    tz = _resolve_output_tz(tz_str)
    addr_list = [a.strip() for a in attendees.split(",")]

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

    if _wants_json(as_json):
        click.echo(to_json_envelope(suggestions, tz=tz))
    else:
        if not suggestions:
            print_error("No available meeting slots found.")
        else:
            console.print(f"[bold cyan]Available slots ({len(suggestions)})[/bold cyan]")
            print_meeting_suggestions(suggestions)


@click.command("people-search")
@click.argument("query")
@click.option("--max", "-n", "max_count", default=10, type=int, help="Max results")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def people_search(query: str, max_count: int, as_json: bool, account_name: str | None):
    """Search people by name for attendee autocomplete."""
    client = _get_client()
    results = client.search_people(query, top=max_count)
    if _wants_json(as_json):
        click.echo(to_json_envelope(results))
    else:
        if not results:
            print_error("No people found.")
        else:
            print_people(results)
