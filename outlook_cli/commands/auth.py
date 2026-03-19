"""Auth commands: login, whoami."""

from __future__ import annotations

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    do_login,
    get_account_name,
    print_error,
    print_success,
    print_whoami,
    to_json_envelope,
    verify_token,
)


@click.command()
@click.option("--force", is_flag=True, help="Force re-login, ignore saved session")
@click.option("--debug", is_flag=True, help="Show debug info about captured requests")
@click.option("--with-token", is_flag=True, help="Read token from standard input instead of using browser")
@account_option
def login(force: bool, debug: bool, with_token: bool, account_name: str | None):
    """Authenticate and cache the bearer token.

    By default, launches a browser to capture the token automatically.
    With --with-token, reads token from stdin (useful with automation/CI).

    Examples:
        outlook login                           # Browser automation
        outlook login --with-token < token.txt   # Read from file
        echo $TOKEN | outlook login --with-token    # Read from pipe
    """
    import sys

    try:
        login_kwargs = {"force": force, "debug": debug}
        if account_name:
            login_kwargs["account_name"] = account_name
        if with_token:
            token = sys.stdin.read().strip()
            if not token:
                print_error("No token provided via stdin.")
                sys.exit(1)
            login_kwargs["token"] = token
        token = do_login(**login_kwargs)
        selected = get_account_name(account_name)
        if verify_token(token):
            print_success(f"Logged in successfully for account '{selected}'. Token cached.")
        else:
            print_error("Login completed but token verification failed.")
    except (RuntimeError, ValueError) as e:
        print_error(str(e))
        sys.exit(1)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def whoami(as_json: bool, account_name: str | None):
    """Show current user info."""
    client = _get_client()
    selected = get_account_name(account_name)
    data = dict(client.get_me())
    data["AccountProfile"] = selected
    if _wants_json(as_json):
        click.echo(to_json_envelope(data))
    else:
        print_whoami(data, account_name=selected)
