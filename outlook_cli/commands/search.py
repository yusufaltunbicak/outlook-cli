"""Search commands: search."""

from __future__ import annotations

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_error,
    print_inbox,
    save_json,
    to_json_envelope,
    print_success,
)


@click.command()
@click.argument("query")
@click.option("--max", "-n", "max_count", default=25, type=int, help="Max results")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def search(query: str, max_count: int, as_json: bool, output: str | None, account_name: str | None):
    """Search messages."""
    client = _get_client()
    messages = client.search_messages(query, top=max_count)

    if _wants_json(as_json):
        if output:
            save_json(messages, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(messages))
    else:
        if not messages:
            print_error("No results found.")
        else:
            print_inbox(messages)
