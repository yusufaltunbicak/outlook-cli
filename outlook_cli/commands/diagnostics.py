"""Offline capabilities and diagnostics: no keychain, auth, or network required."""
from __future__ import annotations

import importlib.metadata
import json
import time
from pathlib import Path

import click

from .. import __version__, account
from ..serialization import to_json_envelope
from ._common import _handle_api_error, account_option


def command_schema(command: click.Command) -> dict:
    parameters = []
    for parameter in command.params:
        row = {"name": parameter.name, "required": parameter.required,
               "type": str(parameter.type), "nargs": parameter.nargs}
        if isinstance(parameter, click.Option):
            row.update(options=parameter.opts + parameter.secondary_opts,
                       multiple=parameter.multiple, flag=parameter.is_flag,
                       help=parameter.help)
        parameters.append(row)
    result = {"name": command.name, "help": command.help, "parameters": parameters}
    if isinstance(command, click.Group):
        result["commands"] = {name: command_schema(cmd) for name, cmd in sorted(command.commands.items())}
    return result


@click.command()
@click.argument("command_name", required=False)
@click.option("--json", "as_json", is_flag=True, help="Output machine-readable command schema")
@_handle_api_error
def schema(command_name: str | None, as_json: bool):
    """Describe available commands, arguments, and output contracts offline."""
    from ..cli import cli
    command = cli
    if command_name:
        for part in command_name.split():
            if not isinstance(command, click.Group) or part not in command.commands:
                raise click.UsageError(f"Unknown command: {command_name}")
            command = command.commands[part]
    click.echo(to_json_envelope({"cli_version": __version__, "schema": command_schema(command),
        "output": {"success": "ok/data/meta", "error": "ok/error", "data_only": "explicit legacy raw export"},
        "references": {"id": "provider ID; prefer for scripts", "display_num": "persistent account-scoped reference, not list row"}}))


@click.command()
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def doctor(as_json: bool, account_name: str | None):
    """Show installation, auth metadata, and supported features without logging in."""
    selected = account.resolve_account_name(account_name)
    paths = account.get_account_paths(selected)
    try:
        installed = importlib.metadata.version("outlook365-cli")
    except importlib.metadata.PackageNotFoundError:
        installed = None
    metadata = {}
    malformed = False
    if paths.token_file.exists():
        try:
            metadata = json.loads(paths.token_file.read_text())
            if not isinstance(metadata, dict):
                metadata, malformed = {}, True
        except (OSError, ValueError):
            malformed = True
    expires = metadata.get("expires_at")
    try:
        expires_in = max(0, int(float(expires) - time.time())) if expires is not None else None
    except (TypeError, ValueError):
        expires_in, malformed = None, True
    click.echo(to_json_envelope({
        "source_version": __version__, "installed_version": installed,
        "version_match": installed == __version__, "source_path": str(Path(__file__).resolve().parents[1]),
        "account": selected,
        "auth": {"metadata_present": paths.token_file.exists(), "metadata_invalid": malformed,
                 "expires_in_seconds": expires_in, "keychain_checked": False},
        "state": {"legacy_id_map_present": paths.id_map_file.exists(),
                  "index_present": (paths.cache_dir / "index.sqlite3").exists(),
                  "local_mail_present": (paths.cache_dir / "mail.sqlite3").exists()},
        "capabilities": {"legacy_rest": True, "graph_read_delta": True, "graph_mutations": False,
                         "local_index": True, "local_mail_store": True, "rest_read_delta": True,
                         "local_research_offline": True, "batch_read": True, "structured_output": True},
        "graph_setup": "Optional: graph-login --client-id APP_ID; requires an Entra public client app and delegated Mail.Read/User.Read consent.",
    }))
