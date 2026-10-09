"""Outlook 365 CLI entry point — command registration only."""

from __future__ import annotations

import sys
import functools
import inspect
import json
import os
from pathlib import Path

import click

from .commands import (
    account as account_mod,
    attachments as attachments_mod,
    auth as auth_mod,
    calendar as calendar_mod,
    categories as categories_mod,
    contacts as contacts_mod,
    folders as folders_mod,
    mail as mail_mod,
    manage as manage_mod,
    open_item as open_item_mod,
    schedule as schedule_mod,
    search as search_mod,
    signatures as signatures_mod,
    summary as summary_mod,
    diagnostics as diagnostics_mod,
    index as index_mod,
    local as local_mod,
)
from .formatter import console

BANNER = r"""
 ╔═╗┬ ┬┌┬┐┬  ┌─┐┌─┐┬┌─  ╔═╗╦  ╦
 ║ ║│ │ │ │  │ ││ │├┴┐  ║  ║  ║
 ╚═╝└─┘ ┴ ┴─┘└─┘└─┘┴ ┴  ╚═╝╩═╝╩
"""

GLOBAL_FLAG_OPTIONS = {"--no-input", "--dry-run", "--profile"}
GLOBAL_VALUE_OPTIONS = {"--enable-commands"}


def _rewrite_global_option_args(args: list[str]) -> list[str]:
    moved: list[str] = []
    remaining: list[str] = []
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--":
            remaining.extend(args[i:])
            break
        if arg in GLOBAL_FLAG_OPTIONS:
            moved.append(arg)
            i += 1
            continue
        matched_value_option = next((opt for opt in GLOBAL_VALUE_OPTIONS if arg == opt or arg.startswith(f"{opt}=")), None)
        if matched_value_option:
            moved.append(arg)
            if arg == matched_value_option and i + 1 < len(args):
                moved.append(args[i + 1])
                i += 2
            else:
                i += 1
            continue
        remaining.append(arg)
        i += 1
    return moved + remaining


def _parse_enabled_commands(value: str | None) -> set[str]:
    if not value:
        return set()
    return {part.strip().lower() for part in value.split(",") if part.strip()}


class OutlookGroup(click.Group):
    """Custom group that shows the Outlook CLI banner in help output."""

    def main(self, args=None, prog_name=None, complete_var=None, standalone_mode=True, windows_expand_args=True, **extra):
        if args is None:
            args = sys.argv[1:]
        args = _rewrite_global_option_args(list(args))
        from .telemetry import start_profile, finish_profile
        profile_run = start_profile() if "--profile" in args else None
        try:
            result = super().main(
                args=args, prog_name=prog_name, complete_var=complete_var,
                standalone_mode=False, windows_expand_args=windows_expand_args, **extra,
            )
            if standalone_mode and isinstance(result, int) and result:
                raise SystemExit(result)
            return result
        except click.ClickException as exc:
            from .serialization import error_json
            if "--json" in args or "--data-only" in args or not sys.stdout.isatty():
                click.echo(error_json("invalid_usage", exc.format_message()))
            else:
                exc.show()
            if standalone_mode:
                raise SystemExit(exc.exit_code)
            raise
        except click.Abort as exc:
            from .serialization import error_json
            interrupted = isinstance(exc.__cause__, KeyboardInterrupt)
            if "--json" in args or not sys.stdout.isatty():
                click.echo(error_json("interrupted" if interrupted else "aborted",
                                      "Operation interrupted; committed checkpoints retained." if interrupted else "Operation aborted."))
            if standalone_mode:
                raise SystemExit(130 if interrupted else 1)
            raise
        finally:
            if profile_run is not None:
                click.echo(json.dumps({"profile": finish_profile(profile_run)}), err=True)

    def format_help(self, ctx, formatter):
        console.print(f"[bold cyan]{BANNER}[/bold cyan]", highlight=False)
        console.print("  [dim]Outlook 365 from your terminal[/dim]")
        console.print()
        super().format_help(ctx, formatter)


@click.group(cls=OutlookGroup, invoke_without_command=True)
@click.version_option(package_name="outlook365-cli")
@click.option("--json", "global_json", is_flag=True, help="Output structured JSON for the selected command")
@click.option("--profile", is_flag=True, help="Write timing/request counters to stderr")
@click.option("--no-input", is_flag=True, help="Never prompt; fail instead (useful for CI)")
@click.option("--dry-run", is_flag=True, help="Do not make changes; print intended actions and exit successfully")
@click.option("--enable-commands", envvar="OUTLOOK_ENABLE_COMMANDS", help="Comma-separated list of enabled top-level commands")
@click.pass_context
def cli(ctx: click.Context, no_input: bool, dry_run: bool, enable_commands: str | None, profile: bool, global_json: bool):
    """Outlook 365 CLI - read, send, and manage emails from the terminal."""
    ctx.ensure_object(dict)
    ctx.obj["json"] = global_json
    ctx.obj["no_input"] = no_input
    ctx.obj["dry_run"] = dry_run
    ctx.obj["enable_commands"] = enable_commands
    if ctx.invoked_subcommand:
        allow = _parse_enabled_commands(enable_commands)
        if allow and "*" not in allow and "all" not in allow:
            command_name = ctx.invoked_subcommand.lower()
            if command_name not in allow:
                raise click.UsageError(
                    f"Command '{command_name}' is not enabled. Use --enable-commands to allow it."
                )
    if ctx.invoked_subcommand is None and not ctx.resilient_parsing:
        click.echo(ctx.get_help())


cli.add_command(diagnostics_mod.schema)
cli.add_command(diagnostics_mod.doctor)
cli.add_command(index_mod.index)
cli.add_command(index_mod.graph_login)
cli.add_command(local_mod.local)

# Auth
cli.add_command(auth_mod.login)
cli.add_command(auth_mod.whoami)
cli.add_command(account_mod.account)

# Mail - Read & Write
cli.add_command(mail_mod.inbox)
cli.add_command(mail_mod.read)
cli.add_command(mail_mod.thread)
cli.add_command(mail_mod.send)
cli.add_command(mail_mod.draft)
cli.add_command(mail_mod.draft_send)
cli.add_command(mail_mod.draft_verify)
cli.add_command(mail_mod.reply)
cli.add_command(mail_mod.reply_draft)
cli.add_command(mail_mod.forward)

# Scheduled send
cli.add_command(schedule_mod.schedule)
cli.add_command(schedule_mod.schedule_list)
cli.add_command(schedule_mod.schedule_cancel)
cli.add_command(schedule_mod.schedule_draft)

# Search
cli.add_command(search_mod.search)
cli.add_command(summary_mod.summary)

# Folders
cli.add_command(folders_mod.folders)
cli.add_command(folders_mod.folder)

# Categories
cli.add_command(categories_mod.categories)
cli.add_command(categories_mod.categorize)
cli.add_command(categories_mod.uncategorize)
cli.add_command(categories_mod.category_rename)
cli.add_command(categories_mod.category_clear)
cli.add_command(categories_mod.category_delete)
cli.add_command(categories_mod.category_create)

# Signatures
cli.add_command(signatures_mod.signature_pull)
cli.add_command(signatures_mod.signature_list)
cli.add_command(signatures_mod.signature_show)
cli.add_command(signatures_mod.signature_delete)

# Management
cli.add_command(manage_mod.mark_read)
cli.add_command(manage_mod.move)
cli.add_command(manage_mod.copy)
cli.add_command(manage_mod.delete)
cli.add_command(manage_mod.flag)
cli.add_command(manage_mod.pin)
cli.add_command(open_item_mod.open_item)

# Attachments
cli.add_command(attachments_mod.attachments)

# Calendar
cli.add_command(calendar_mod.calendar)
cli.add_command(calendar_mod.event)
cli.add_command(calendar_mod.event_create)
cli.add_command(calendar_mod.event_update)
cli.add_command(calendar_mod.event_delete)
cli.add_command(calendar_mod.event_instances)
cli.add_command(calendar_mod.event_respond)
cli.add_command(calendar_mod.calendars_cmd)
cli.add_command(calendar_mod.free_busy)
cli.add_command(calendar_mod.people_search)

# Contacts
cli.add_command(contacts_mod.contacts)


# Common output options are installed from this single registry so commands cannot
# drift into incompatible stdout/file contracts. Unknown options remain parser errors.
def install_output_contract(command, path=""):
    if isinstance(command, click.Group):
        for name, child in command.commands.items():
            install_output_contract(child, (path + " " + name).strip())
        return
    if getattr(command, "_output_contract_installed", False):
        return
    command._output_contract_installed = True
    callback = command.callback
    accepted = set(inspect.signature(callback).parameters)
    existing = {p.name for p in command.params}
    additions = [
        click.Option(["--json", "as_json"], is_flag=True, help="Output structured JSON"),
        click.Option(["--output", "-o"], type=click.Path(dir_okay=False), help="Write the same JSON result to a file"),
        click.Option(["--data-only"], is_flag=True, help="Export raw data without the JSON envelope (legacy format)"),
        click.Option(["--view"], type=click.Choice(["full", "compact"]), default="full", help="JSON message detail level"),
        click.Option(["--fields"], help="Comma-separated JSON fields"),
    ]
    mail_read = path in {"inbox", "search", "folder", "read", "thread"}
    if path in {"send", "draft", "schedule", "forward"}:
        additions.append(click.Option(["--to", "extra_to"], multiple=True, help="Additional TO recipients (repeatable, comma or semicolon separated)"))
    if mail_read:
        additions.append(click.Option(["--body-format", "--body", "body_format"], type=click.Choice(["none", "preview", "text", "html"]), help="Message body representation"))
    if path in {"inbox", "search", "folder", "calendar", "contacts", "event-instances"}:
        additions.append(click.Option(["--all", "all_pages"], is_flag=True, help="Follow available pages; consult completeness metadata for provider limits"))
    for option in additions:
        if option.name not in existing:
            command.params.append(option)
    for param in command.params:
        if isinstance(param, click.Option) and "--max" in param.opts:
            if "--limit" not in param.opts:
                param.opts.append("--limit")
            if param.type is click.INT:
                param.type = click.IntRange(min=1)

    @functools.wraps(callback)
    def wrapped(**kwargs):
        from .commands._common import _wants_json, maybe_dry_run
        from .serialization import MAIL_API_FIELDS, to_json_envelope
        ctx = click.get_current_context()
        fields = [field.strip() for field in (kwargs.get("fields") or "").split(",") if field.strip()]
        if mail_read and fields:
            unknown = set(fields) - set(MAIL_API_FIELDS)
            if unknown:
                raise click.BadParameter("Unknown fields: " + ", ".join(sorted(unknown)), param_hint="--fields")
        output = kwargs.get("output")
        if output:
            target = Path(output)
            if target.is_dir() or not target.parent.is_dir() or not os.access(target.parent, os.W_OK) or (target.exists() and not os.access(target, os.W_OK)):
                raise click.BadParameter("Output path must be writable and its parent directory must exist", param_hint="--output")
        ctx._outlook_output = {key: kwargs.get(key) for key in ("output", "data_only", "view", "body_format", "all_pages")}
        ctx._outlook_output["fields"] = fields
        ctx._outlook_emitted = False
        ctx._outlook_extra_to = kwargs.get("extra_to", ())
        json_mode = _wants_json(bool(kwargs.get("as_json"))) or bool(kwargs.get("output") or kwargs.get("data_only") or fields or kwargs.get("view") == "compact" or kwargs.get("body_format"))
        ctx._outlook_json_mode = json_mode
        if "as_json" in accepted:
            kwargs["as_json"] = json_mode
        # Guards for mutation paths predating the shared helper. Other commands
        # validate and normalize first, then call maybe_dry_run themselves.
        if path in {"login"}:
            maybe_dry_run(path, {k: v for k, v in kwargs.items() if k in accepted and k not in {"as_json", "account_name"}})
        result = callback(**{k: v for k, v in kwargs.items() if k in accepted})
        if json_mode and not ctx._outlook_emitted:
            payload = result if result is not None else {"status": "completed", "operation": path}
            click.echo(to_json_envelope(payload))
        return result

    from .commands._common import _handle_api_error
    command.callback = _handle_api_error(wrapped)


install_output_contract(cli)
