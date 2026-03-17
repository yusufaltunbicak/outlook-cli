"""Mail commands: inbox, read, thread, send, draft, draft-send, reply, reply-draft, forward."""

from __future__ import annotations

import os

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    console,
    print_email,
    print_email_raw,
    print_inbox,
    print_success,
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
            print_inbox(messages)


@click.command()
@click.argument("message_id")
@click.option("--raw", is_flag=True, help="Show raw HTML body")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def read(message_id: str, raw: bool, as_json: bool, account_name: str | None):
    """Read an email by its display number."""
    client = _get_client()
    email = client.get_message(message_id)

    if _wants_json(as_json):
        click.echo(to_json_envelope(email))
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


@click.command()
@click.argument("message_id")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def thread(message_id: str, as_json: bool, account_name: str | None):
    """Show the full conversation thread for a message."""
    from ..formatter import print_thread

    client = _get_client()
    messages = client.get_thread(message_id)

    if _wants_json(as_json):
        click.echo(to_json_envelope(messages))
    else:
        if len(messages) <= 1:
            print_success("This message is not part of a conversation thread.")
            if messages:
                print_email(messages[0])
        else:
            print_thread(messages)


@click.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body")
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def send(to: str, subject: str, body: str, cc: tuple, attach: tuple, is_html: bool, sig_name: str | None, as_json: bool, yes: bool, account_name: str | None):
    """Send an email. TO can be comma-separated for multiple recipients."""
    from ..signature_manager import append_signature, get_signature

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
        _show_attachment_info(attach)
        click.confirm("Send this email?", abort=True)

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
@click.argument("body")
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def draft(to: str, subject: str, body: str, cc: tuple, attach: tuple, is_html: bool, sig_name: str | None, as_json: bool, account_name: str | None):
    """Create a draft email without sending. TO can be comma-separated."""
    from ..signature_manager import append_signature, get_signature

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    client = _get_client()
    to_list = [addr.strip() for addr in to.split(",")]
    cc_list = list(cc) if cc else None
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


@click.command()
@click.argument("message_id")
@click.argument("body")
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def reply(message_id: str, body: str, reply_all: bool, attach: tuple, yes: bool, account_name: str | None):
    """Reply to an email."""
    client = _get_client()
    if not yes:
        action = "Reply all" if reply_all else "Reply"
        console.print(f"  [bold]{action} to #{message_id}[/bold]")
        console.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        _show_attachment_info(attach)
        click.confirm("Send this reply?", abort=True)

    if attach:
        # Draft flow: create reply draft -> attach -> send
        draft_email = client.create_reply_draft(message_id, comment=body, reply_all=reply_all)
        client.attach_files(draft_email.id, list(attach))
        client.send_draft(draft_email.id)
    else:
        client.reply(message_id, body, reply_all=reply_all)

    action = "Reply all" if reply_all else "Reply"
    print_success(f"{action} sent for message #{message_id}")


@click.command(name="reply-draft")
@click.argument("message_id")
@click.argument("body", default="")
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--html", "is_html", is_flag=True, help="Body is HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def reply_draft(message_id: str, body: str, reply_all: bool, attach: tuple, is_html: bool, sig_name: str | None, as_json: bool, account_name: str | None):
    """Create a reply draft without sending."""
    from ..signature_manager import append_signature, get_signature

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name and body:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

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
    to_list = [addr.strip() for addr in to.split(",")]
    if not yes:
        console.print(f"  [bold]Forward #{message_id} to:[/bold] {', '.join(to_list)}")
        if comment:
            console.print(f"  [bold]Comment:[/bold] {comment[:100]}{'...' if len(comment) > 100 else ''}")
        _show_attachment_info(attach)
        click.confirm("Forward this email?", abort=True)

    client = _get_client()

    if attach:
        # Draft flow: create forward draft -> attach -> send
        draft_email = client.create_forward_draft(message_id, to_list, comment=comment)
        client.attach_files(draft_email.id, list(attach))
        client.send_draft(draft_email.id)
    else:
        client.forward(message_id, to_list, comment=comment)

    print_success(f"Message #{message_id} forwarded to {to}")
