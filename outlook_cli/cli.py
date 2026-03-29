"""Outlook 365 CLI entry point — command registration only."""

from __future__ import annotations

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


class OutlookGroup(click.Group):
    """Custom group that shows the Outlook CLI banner in help output."""

    def format_help(self, ctx, formatter):
        console.print(f"[bold cyan]{BANNER}[/bold cyan]", highlight=False)
        console.print("  [dim]Outlook 365 from your terminal[/dim]")
        console.print()
        super().format_help(ctx, formatter)


@click.group(cls=OutlookGroup, invoke_without_command=True)
@click.version_option(package_name="outlook365-cli")
@click.pass_context
def cli(ctx: click.Context):
    """Outlook 365 CLI - read, send, and manage emails from the terminal."""
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
