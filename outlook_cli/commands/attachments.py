"""Metadata-first attachment listing and safe downloads, independent of output mode."""
from __future__ import annotations

import base64
import os
from dataclasses import asdict
from pathlib import Path

import click

from ._common import _get_client, _handle_api_error, _wants_json, account_option, maybe_dry_run, print_attachments, print_success, to_json_envelope
from ..exceptions import error_code_for_exception


def _safe_name(name: str) -> str:
    # Outlook attachment names are untrusted and may use Windows separators.
    clean = str(name).replace("\\", "/").rsplit("/", 1)[-1].strip()
    if clean in {"", ".", ".."} or any(ord(char) < 32 for char in clean):
        raise ValueError("Attachment has an unsafe or empty filename")
    return clean


@click.command()
@click.argument("message_id")
@click.option("-d", "--download", is_flag=True, help="Download all attachments without overwriting existing files")
@click.option("--save-to", type=click.Path(file_okay=False), default=".", help="Download directory")
@click.option("--include-content", is_flag=True, help="Include base64 attachment contents in JSON (large output)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def attachments(message_id: str, download: bool, save_to: str, include_content: bool, as_json: bool, account_name: str | None):
    """List metadata or download attachments, including when stdout is piped."""
    if download:
        maybe_dry_run("attachments.download", {"message_id": message_id, "save_to": save_to})
    client = _get_client()
    atts = client.get_attachments(message_id)
    results = []
    failed = 0
    save_path = Path(save_to)
    if download:
        save_path.mkdir(parents=True, exist_ok=True)
    for att in atts:
        row = asdict(att)
        row.pop("content_bytes", None)
        row["status"] = "listed"
        try:
            if download or include_content:
                full = att if att.content_bytes is not None else client.download_attachment(message_id, att.id)
                content = full.content_bytes
                if content is None:
                    raise ValueError("Attachment has no downloadable file content")
                if download:
                    target = save_path / _safe_name(att.name)
                    contents = base64.b64decode(content, validate=True)
                    # O_EXCL also rejects symlinks and duplicate names within a mail.
                    fd = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
                    try:
                        with os.fdopen(fd, "wb") as stream:
                            stream.write(contents)
                    except BaseException:
                        target.unlink(missing_ok=True)
                        raise
                    row.update(status="downloaded", path=str(target.resolve()), bytes_written=len(contents))
                    print_success(f"Saved: {target}")
                if include_content:
                    row["content_bytes"] = content
        except Exception as exc:
            failed += 1
            row.update(status="failed", error={"code": error_code_for_exception(exc), "message": str(exc)})
        results.append(row)
    if _wants_json(as_json):
        click.echo(to_json_envelope(results, ok=not failed, meta={"returned": len(results), "failed": failed, "partial": bool(failed and failed < len(results))}, error={"code": "partial_failure", "message": f"{failed} attachment(s) failed"} if failed else None))
    elif atts:
        print_attachments(atts)
        for row in results:
            if row["status"] == "failed":
                click.echo(f"{row['name']}: {row['error']['message']}", err=True)
    else:
        print_success("No attachments.")
    if failed:
        raise click.exceptions.Exit(1)
