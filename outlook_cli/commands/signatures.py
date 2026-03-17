"""Signature commands: signature-pull, signature-list, signature-show, signature-delete."""

from __future__ import annotations

import click

from ._common import (
    _handle_api_error,
    account_option,
    cfg,
    console,
    get_token,
    print_success,
)


@click.command("signature-pull")
@click.option("--name", "-n", default=None, help="Name for the signature (default: auto-detect)")
@account_option
@_handle_api_error
def signature_pull(name: str | None, account_name: str | None):
    """Extract your signature from a recent sent email and save it."""
    from ..signature_manager import pull_signature, save_signature

    token = get_token()
    sig_html, source_subject = pull_signature(token)

    if not name:
        name = click.prompt("Signature name", default="default")

    path = save_signature(name, sig_html)
    print_success(f"Signature '{name}' saved from: {source_subject}")
    console.print(f"  [dim]{path}[/dim]")


@click.command("signature-list")
@account_option
def signature_list(account_name: str | None):
    """List saved signatures."""
    from ..signature_manager import list_signatures

    sigs = list_signatures()
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

    sig_html = get_signature(name)
    text = BeautifulSoup(sig_html, "html.parser").get_text("\n", strip=True)
    console.print(text)


@click.command("signature-delete")
@click.argument("name")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def signature_delete(name: str, yes: bool, account_name: str | None):
    """Delete a saved signature."""
    from ..signature_manager import delete_signature

    if not yes:
        click.confirm(f"Delete signature '{name}'?", abort=True)
    delete_signature(name)
    print_success(f"Deleted signature '{name}'")
