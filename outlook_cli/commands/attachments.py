"""Attachment commands: attachments."""

from __future__ import annotations

import base64
from pathlib import Path

import click

from ._common import (
    _get_client,
    _handle_api_error,
    print_attachments,
    print_success,
    to_json_envelope,
)


@click.command()
@click.argument("message_id")
@click.option("-d", "--download", is_flag=True, help="Download all attachments")
@click.option("--save-to", type=click.Path(), default=".", help="Download directory")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def attachments(message_id: str, download: bool, save_to: str, as_json: bool):
    """List or download attachments for a message."""
    client = _get_client()
    atts = client.get_attachments(message_id)

    if not atts:
        print_success("No attachments.")
        return

    if as_json:
        click.echo(to_json_envelope(atts))
        return

    print_attachments(atts)

    if download:
        save_path = Path(save_to)
        save_path.mkdir(parents=True, exist_ok=True)
        for att in atts:
            if att.content_bytes:
                file_path = save_path / att.name
                file_path.write_bytes(base64.b64decode(att.content_bytes))
                print_success(f"  Saved: {file_path}")
            else:
                full = client.download_attachment(message_id, att.id)
                if full.content_bytes:
                    file_path = save_path / full.name
                    file_path.write_bytes(base64.b64decode(full.content_bytes))
                    print_success(f"  Saved: {file_path}")
