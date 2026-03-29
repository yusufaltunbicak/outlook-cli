"""Open messages and events in Outlook on the web."""

from __future__ import annotations

import webbrowser

import click

from ..exceptions import OutlookCliError
from ._common import _get_client, _handle_api_error, account_option, print_success


@click.command("open")
@click.argument("item_id")
@click.option("--print-url", is_flag=True, help="Print the Outlook web URL instead of opening a browser")
@account_option
@_handle_api_error
def open_item(item_id: str, print_url: bool, account_name: str | None):
    """Open a message or event in Outlook on the web."""
    client = _get_client(account_name)
    kind, url = client.get_open_target(item_id)

    if print_url:
        click.echo(url)
        return

    if not webbrowser.open(url):
        raise OutlookCliError(f"Could not open a browser automatically. URL: {url}")

    label = f"#{item_id}" if item_id.isdigit() else item_id
    print_success(f"Opened {kind} {label} in browser")
