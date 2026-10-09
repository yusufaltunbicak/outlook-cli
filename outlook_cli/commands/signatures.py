"""Signature commands: signature-pull, signature-list, signature-show, signature-delete."""

from __future__ import annotations

import click
import sys
from contextlib import redirect_stdout

from ._common import (
    _handle_api_error,
    _get_client,
    _wants_json,
    is_no_input_mode,
    to_json_envelope,
    account_option,
    confirm_action,
    cfg,
    console,
    get_token,
    get_account_name,
    maybe_dry_run,
    print_success,
)


@click.command("signature-pull")
@click.option("--name", "-n", default=None, help="Name for the signature (default: auto-detect)")
@account_option
@_handle_api_error
def signature_pull(name: str | None, account_name: str | None):
    """Extract your signature from a recent sent email and save it."""
    from ..signature_manager import pull_signature, save_signature

    from ..category_manager import bind_client
    maybe_dry_run("signature-pull", {"name": name})
    client = _get_client()
    bind_client(client)
    token = client._token
    sig_html, source_subject = pull_signature(token)

    if not name:
        if is_no_input_mode():
            name = "default"
        else:
            with redirect_stdout(sys.stderr):
                name = click.prompt("Signature name", default="default", err=True)

    path = save_signature(name, sig_html, account_name=get_account_name(account_name))
    print_success(f"Signature '{name}' saved from: {source_subject}")
    console.print(f"  [dim]{path}[/dim]")
    return {"name": name, "path": str(path), "source_subject": source_subject}


@click.command("signature-list")
@account_option
def signature_list(account_name: str | None):
    """List saved signatures."""
    from ..signature_manager import list_signatures

    sigs = list_signatures(account_name=get_account_name(account_name))
    if _wants_json(False):
        click.echo(to_json_envelope(sigs))
        return
    if not sigs:
        print_success("No signatures saved. Run 'outlook signature-pull' to extract one.")
    else:
        for s in sigs:
            default = " [bold cyan](default)[/bold cyan]" if s == cfg.get("default_signature") else ""
            console.print(f"  {s}{default}")


@click.command("signature-show")
@click.argument("name")
@account_option
@_handle_api_error
def signature_show(name: str, account_name: str | None):
    """Preview a saved signature."""
    from ..signature_manager import get_signature

    from bs4 import BeautifulSoup

    sig_html = get_signature(name, account_name=get_account_name(account_name))
    text = BeautifulSoup(sig_html, "html.parser").get_text("\n", strip=True)
    if _wants_json(False):
        click.echo(to_json_envelope({"name": name, "html": sig_html, "text": text}))
    else:
        console.print(text)


@click.command("signature-delete")
@click.argument("name")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def signature_delete(name: str, yes: bool, account_name: str | None):
    """Delete a saved signature."""
    from ..signature_manager import delete_signature

    maybe_dry_run("signature-delete", {"name": name})
    if not yes:
        confirm_action(f"Delete signature '{name}'?", action=f"delete signature '{name}'")
    delete_signature(name, account_name=get_account_name(account_name))
    print_success(f"Deleted signature '{name}'")
