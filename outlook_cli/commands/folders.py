"""Folder commands: folders, folder."""

from __future__ import annotations

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    console,
    get_category_color_map,
    print_folders,
    print_inbox,
    print_success,
    save_json,
    to_json_envelope,
)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def folders(as_json: bool, output: str | None, account_name: str | None):
    """List mail folders."""
    client = _get_client()
    folder_list = client.get_folders()

    if _wants_json(as_json):
        if output:
            save_json(folder_list, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(folder_list))
    else:
        print_folders(folder_list)


@click.command()
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
@account_option
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
    account_name: str | None,
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

    if _wants_json(as_json):
        click.echo(to_json_envelope(messages))
    else:
        if not messages:
            print_success(f"No messages found in '{name}'.")
        else:
            console.print(f"[bold cyan]Folder: {name}[/bold cyan]")
            print_inbox(messages, category_colors=get_category_color_map(client, messages))
