"""Contact commands: contacts."""

from __future__ import annotations

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_contacts,
    print_success,
    save_json,
    to_json_envelope,
)


@click.command()
@click.option("--max", "-n", "max_count", default=50, type=int, help="Max contacts")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def contacts(max_count: int, as_json: bool, output: str | None, account_name: str | None):
    """List contacts."""
    client = _get_client()
    contact_list = client.get_contacts(top=max_count)

    if _wants_json(as_json):
        if output:
            save_json(contact_list, output)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(contact_list))
    else:
        if not contact_list:
            print_success("No contacts found.")
        else:
            print_contacts(contact_list)
