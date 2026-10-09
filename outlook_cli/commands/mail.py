"""Mail commands: inbox, read, thread, send, draft, draft-send, reply, reply-draft, forward."""

from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor

from ..recipients import normalize_recipients
from ..exceptions import error_code_for_exception
from ..serialization import message_fetch_options

import click

from ._common import (
    _get_client,
    is_dry_run_mode,
    _handle_api_error,
    _wants_json,
    account_option,
    confirm_action,
    cfg,
    console,
    get_category_color_map,
    get_account_name,
    maybe_dry_run,
    print_email,
    print_email_raw,
    print_inbox,
    print_success,
    resolve_body_input,
    save_json,
    to_json_envelope,
)


def _format_file_size(size: int) -> str:
    """Human-readable file size."""
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def _show_attachment_info(file_paths: tuple[str, ...]) -> None:
    """Print attachment info in confirmation prompt."""
    if not file_paths:
        return
    parts = []
    for fp in file_paths:
        name = os.path.basename(fp)
        size = os.path.getsize(fp)
        parts.append(f"{name} ({_format_file_size(size)})")
    console.print(f"  [bold]Attachments:[/bold] {', '.join(parts)}")


@click.command()
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
@account_option
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
    account_name: str | None,
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
        **message_fetch_options(),
    )

    if _wants_json(as_json):
        if output:
            save_json(messages, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(messages))
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
            print_inbox(messages, category_colors=get_category_color_map(client, messages))


def _bulk_read(client, message_ids, *, thread_mode=False, peek=False, workers=4):
    def fetch(mid):
        try:
            data = client.get_thread(mid) if thread_mode else client.get_message(mid)
            row = {"id": mid, "ok": True, "data": data}
            if not thread_mode and not peek and not data.is_read:
                try:
                    client.mark_read(mid)
                except Exception as exc:
                    row["ok"] = False
                    row["error"] = {"code": error_code_for_exception(exc), "message": "Message fetched but marking read failed: " + str(exc)}
            if getattr(data, "meta", None):
                row["meta"] = data.meta
            return row
        except Exception as exc:
            return {"id": mid, "ok": False, "error": {"code": error_code_for_exception(exc), "message": str(exc)}}
    with ThreadPoolExecutor(max_workers=min(workers, len(message_ids))) as pool:
        return list(pool.map(fetch, message_ids))


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--raw", is_flag=True, help="Show raw HTML body")
@click.option("--peek", is_flag=True, help="Read without marking the message read")
@click.option("--workers", type=click.IntRange(1, 4), default=4, help="Maximum concurrent reads (1-4)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def read(message_ids: tuple, raw: bool, peek: bool, workers: int, as_json: bool, account_name: str | None):
    """Read one or more real IDs or persistent display numbers, in input order."""
    if is_dry_run_mode():
        peek = True
    client = _get_client()
    results = _bulk_read(client, message_ids, peek=peek, workers=workers)
    failed = sum(not row["ok"] for row in results)
    if _wants_json(as_json):
        single = len(results) == 1 and not failed
        data = results[0]["data"] if single else results
        click.echo(to_json_envelope(data, ok=not failed, meta={"returned": len(results), "failed": failed, "partial": bool(failed and failed < len(results))}, error={"code": "partial_failure", "message": f"{failed} message operation(s) failed"} if failed else None))
    else:
        for row in results:
            if "data" in row:
                (print_email_raw if raw else print_email)(row["data"])
            if not row["ok"]:
                click.echo(f"{row['id']}: {row['error']['message']}", err=True)
    if failed:
        raise click.exceptions.Exit(1)


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--workers", type=click.IntRange(1, 4), default=4, help="Maximum concurrent thread reads (1-4)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def thread(message_ids: tuple, workers: int, as_json: bool, account_name: str | None):
    """Read one or more conversations; completeness is reported in metadata."""
    from ..formatter import print_thread
    client = _get_client()
    results = _bulk_read(client, message_ids, thread_mode=True, workers=workers)
    failed = sum(not row["ok"] for row in results)
    if _wants_json(as_json):
        single = len(results) == 1 and not failed
        data = results[0]["data"] if single else results
        click.echo(to_json_envelope(data, ok=not failed, meta={"failed": failed, "partial": bool(failed and failed < len(results))}, error={"code": "partial_failure", "message": f"{failed} thread(s) failed"} if failed else None))
    else:
        for row in results:
            if row["ok"]:
                messages = row["data"]
                if len(messages) <= 1:
                    print_success("This message is not part of a conversation thread.")
                    if messages:
                        print_email(messages[0])
                else:
                    print_thread(messages)
            else:
                click.echo(f"{row['id']}: {row['error']['message']}", err=True)
    if failed:
        raise click.exceptions.Exit(1)


@click.command("draft-verify")
@click.argument("message_id")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def draft_verify(message_id: str, as_json: bool, account_name: str | None):
    """Read actual recipients and attachment metadata for an existing draft."""
    from dataclasses import asdict
    client = _get_client()
    email = client.get_message(message_id)
    attachments = []
    for attachment in client.get_attachments(message_id):
        metadata = asdict(attachment)
        metadata.pop("content_bytes", None)
        attachments.append(metadata)
    data = {"id": email.id, "subject": email.subject, "to": email.to, "cc": email.cc, "attachments": attachments}
    click.echo(to_json_envelope(data))


@click.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body", required=False)
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def send(to: str, subject: str, body: str | None, cc: tuple, attach: tuple, body_file: str | None, is_html: bool, sig_name: str | None, as_json: bool, yes: bool, account_name: str | None):
    """Send an email. TO can be comma-separated for multiple recipients."""
    from ..signature_manager import append_signature, get_signature

    body = resolve_body_input(body, body_file)
    if not body:
        raise click.UsageError("Provide BODY or --body-file.")

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name, account_name=get_account_name(account_name))
        body, is_html = append_signature(body, sig_html, is_html)

    to_list = normalize_recipients([to, *getattr(click.get_current_context(), "_outlook_extra_to", ())], field="TO", required=True)
    cc_list = normalize_recipients(cc, field="CC") or None
    maybe_dry_run(
        "send",
        {
            "to": to_list,
            "subject": subject,
            "body": body,
            "cc": cc_list,
            "attach": list(attach),
            "html": is_html,
        },
    )

    if not yes:
        console.print(f"  [bold]To:[/bold] {', '.join(to_list)}")
        if cc_list:
            console.print(f"  [bold]CC:[/bold] {', '.join(cc_list)}")
        console.print(f"  [bold]Subject:[/bold] {subject}")
        console.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        _show_attachment_info(attach)
        confirm_action("Send this email?", action="send this email")

    client = _get_client()

    if attach:
        # Draft flow: create draft -> attach files -> send
        email = client.create_draft(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)
        client.attach_files(email.id, list(attach))
        client.send_draft(email.id)
    else:
        client.send_mail(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)

    if _wants_json(as_json):
        click.echo(to_json_envelope({"status": "sent", "to": to_list, "subject": subject}))
    else:
        print_success(f"Mail sent to {to}")


@click.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body", required=False)
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def draft(to: str, subject: str, body: str | None, cc: tuple, attach: tuple, body_file: str | None, is_html: bool, sig_name: str | None, as_json: bool, account_name: str | None):
    """Create a draft email without sending. TO can be comma-separated."""
    from ..signature_manager import append_signature, get_signature

    body = resolve_body_input(body, body_file)
    if not body:
        raise click.UsageError("Provide BODY or --body-file.")

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name, account_name=get_account_name(account_name))
        body, is_html = append_signature(body, sig_html, is_html)

    to_list = normalize_recipients([to, *getattr(click.get_current_context(), "_outlook_extra_to", ())], field="TO", required=True)
    cc_list = normalize_recipients(cc, field="CC") or None
    maybe_dry_run(
        "draft",
        {
            "to": to_list,
            "subject": subject,
            "body": body,
            "cc": cc_list,
            "attach": list(attach),
            "html": is_html,
        },
    )
    client = _get_client()
    email = client.create_draft(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)

    if attach:
        client.attach_files(email.id, list(attach))

    if _wants_json(as_json):
        click.echo(to_json_envelope(email))
    else:
        print_success(f"Draft created: {subject} (to: {to})")


@click.command(name="draft-send")
@click.argument("message_id")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def draft_send(message_id: str, yes: bool, account_name: str | None):
    """Send an existing draft by its message number."""
    maybe_dry_run("draft-send", {"message_id": message_id})
    client = _get_client()
    if not yes:
        email = client.get_message(message_id)
        console.print(f"  [bold]To:[/bold] {', '.join(r.address for r in email.to)}")
        if email.cc:
            console.print(f"  [bold]CC:[/bold] {', '.join(r.address for r in email.cc)}")
        console.print(f"  [bold]Subject:[/bold] {email.subject}")
        confirm_action(f"Send draft #{message_id}?", action=f"send draft #{message_id}")
    client.send_draft(message_id)
    print_success(f"Draft #{message_id} sent")
    return {"status": "sent", "id": message_id}


@click.command()
@click.argument("message_id")
@click.argument("body", required=False)
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def reply(message_id: str, body: str | None, reply_all: bool, attach: tuple, body_file: str | None, yes: bool, account_name: str | None):
    """Reply to an email."""
    body = resolve_body_input(body, body_file)
    if not body:
        raise click.UsageError("Provide BODY or --body-file.")
    maybe_dry_run(
        "reply",
        {
            "message_id": message_id,
            "body": body,
            "reply_all": reply_all,
            "attach": list(attach),
        },
    )
    client = _get_client()
    if not yes:
        action = "Reply all" if reply_all else "Reply"
        console.print(f"  [bold]{action} to #{message_id}[/bold]")
        console.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        _show_attachment_info(attach)
        confirm_action("Send this reply?", action="send this reply")

    if attach:
        # Draft flow: create reply draft -> attach -> send
        draft_email = client.create_reply_draft(message_id, comment=body, reply_all=reply_all)
        client.attach_files(draft_email.id, list(attach))
        client.send_draft(draft_email.id)
    else:
        client.reply(message_id, body, reply_all=reply_all)

    action = "Reply all" if reply_all else "Reply"
    print_success(f"{action} sent for message #{message_id}")
    return {"status": "sent", "id": message_id, "reply_all": reply_all}


@click.command(name="reply-draft")
@click.argument("message_id")
@click.argument("body", required=False)
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--html", "is_html", is_flag=True, help="Body is HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def reply_draft(message_id: str, body: str | None, reply_all: bool, attach: tuple, body_file: str | None, is_html: bool, sig_name: str | None, as_json: bool, account_name: str | None):
    """Create a reply draft without sending."""
    from ..signature_manager import append_signature, get_signature

    body = resolve_body_input(body, body_file)
    sig_name = sig_name or cfg.get("default_signature")
    if sig_name and body:
        sig_html = get_signature(sig_name, account_name=get_account_name(account_name))
        body, is_html = append_signature(body, sig_html, is_html)

    maybe_dry_run(
        "reply-draft",
        {
            "message_id": message_id,
            "body": body,
            "reply_all": reply_all,
            "attach": list(attach),
            "html": is_html,
        },
    )
    client = _get_client()
    email = client.create_reply_draft(message_id, comment=body, reply_all=reply_all, html=is_html)

    if attach:
        client.attach_files(email.id, list(attach))

    action = "Reply-all" if reply_all else "Reply"
    if _wants_json(as_json):
        click.echo(to_json_envelope(email))
    else:
        print_success(f"{action} draft created for message #{message_id}")


@click.command()
@click.argument("message_id")
@click.argument("to")
@click.option("--comment", "-c", default="", help="Add a comment to the forwarded message")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def forward(message_id: str, to: str, comment: str, attach: tuple, yes: bool, account_name: str | None):
    """Forward an email."""
    to_list = normalize_recipients([to, *getattr(click.get_current_context(), "_outlook_extra_to", ())], field="TO", required=True)
    maybe_dry_run(
        "forward",
        {
            "message_id": message_id,
            "to": to_list,
            "comment": comment,
            "attach": list(attach),
        },
    )
    if not yes:
        console.print(f"  [bold]Forward #{message_id} to:[/bold] {', '.join(to_list)}")
        if comment:
            console.print(f"  [bold]Comment:[/bold] {comment[:100]}{'...' if len(comment) > 100 else ''}")
        _show_attachment_info(attach)
        confirm_action("Forward this email?", action="forward this email")

    client = _get_client()

    if attach:
        # Draft flow: create forward draft -> attach -> send
        draft_email = client.create_forward_draft(message_id, to_list, comment=comment)
        client.attach_files(draft_email.id, list(attach))
        client.send_draft(draft_email.id)
    else:
        client.forward(message_id, to_list, comment=comment)

    print_success(f"Message #{message_id} forwarded to {to}")
    return {"status": "forwarded", "id": message_id, "to": to_list}
