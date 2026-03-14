"""Auth commands: login, whoami."""

from __future__ import annotations

import click

from ._common import (
    _get_client,
    _handle_api_error,
    do_login,
    print_error,
    print_success,
    print_whoami,
    to_json_envelope,
    verify_token,
)


@click.command()
@click.option("--force", is_flag=True, help="Force re-login, ignore saved session")
@click.option("--debug", is_flag=True, help="Show debug info about captured requests")
def login(force: bool, debug: bool):
    """Authenticate via browser and cache the token."""
    try:
        token = do_login(force=force, debug=debug)
        if verify_token(token):
            print_success("Logged in successfully. Token cached.")
        else:
            print_error("Login completed but token verification failed.")
    except RuntimeError as e:
        print_error(str(e))
        import sys
        sys.exit(1)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def whoami(as_json: bool):
    """Show current user info."""
    client = _get_client()
    data = client.get_me()
    if as_json:
        click.echo(to_json_envelope(data))
    else:
        print_whoami(data)
