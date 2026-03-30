"""Outlook 365 CLI entry point — command registration only."""

from __future__ import annotations

import sys

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
)
from .formatter import console

BANNER = r"""
 ╔═╗┬ ┬┌┬┐┬  ┌─┐┌─┐┬┌─  ╔═╗╦  ╦
 ║ ║│ │ │ │  │ ││ │├┴┐  ║  ║  ║
 ╚═╝└─┘ ┴ ┴─┘└─┘└─┘┴ ┴  ╚═╝╩═╝╩
"""

GLOBAL_FLAG_OPTIONS = {"--no-input", "--dry-run"}
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
        return super().main(
            args=args,
            prog_name=prog_name,
            complete_var=complete_var,
            standalone_mode=standalone_mode,
            windows_expand_args=windows_expand_args,
            **extra,
        )

    def format_help(self, ctx, formatter):
        console.print(f"[bold cyan]{BANNER}[/bold cyan]", highlight=False)
        console.print("  [dim]Outlook 365 from your terminal[/dim]")
        console.print()
        super().format_help(ctx, formatter)


@click.group(cls=OutlookGroup, invoke_without_command=True)
@click.version_option(package_name="outlook365-cli")
@click.option("--no-input", is_flag=True, help="Never prompt; fail instead (useful for CI)")
@click.option("--dry-run", is_flag=True, help="Do not make changes; print intended actions and exit successfully")
@click.option("--enable-commands", envvar="OUTLOOK_ENABLE_COMMANDS", help="Comma-separated list of enabled top-level commands")
@click.pass_context
def cli(ctx: click.Context, no_input: bool, dry_run: bool, enable_commands: str | None):
    """Outlook 365 CLI - read, send, and manage emails from the terminal."""
    ctx.ensure_object(dict)
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
