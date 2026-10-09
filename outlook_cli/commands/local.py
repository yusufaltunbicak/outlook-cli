"""Local-first mail research. Every read here works without authentication."""
from __future__ import annotations

import time
from contextlib import ExitStack
from pathlib import Path
from urllib.parse import quote

import click
import httpx

from .. import account
from ..exceptions import OutlookCliError, RateLimitError, error_code_for_exception
from ..graph import GraphReader
from ..locking import file_lock
from ..mail_store import MailStore
from ..mail_sync import GraphSyncReader, RestReader, sync_folder
from ..pagination import Page, validated_next_link
from ..serialization import to_json_envelope
from ._common import (
    _get_client,
    _handle_api_error,
    account_option,
    confirm_action,
    maybe_dry_run,
)


def store_path(profile):
    return account.get_account_paths(profile).cache_dir / "mail.sqlite3"


def _store(account_name):
    return MailStore(store_path(account.resolve_account_name(account_name)))


def discover_folders(reader):
    graph = reader.backend == "graph"
    root = "/me/mailFolders" if graph else "/MailFolders"
    queue, found, seen = [root], [], set()
    while queue:
        path, params = queue.pop(0), {"$top": 100}
        if graph:
            params["includeHiddenFolders"] = "true"
        links = set()
        for _ in range(1000):
            data = reader.get(path, params=params)
            if not isinstance(data.get("value"), list):
                raise OutlookCliError("Incomplete folder inventory; store was not reconciled.")
            for item in data["value"]:
                identity = item.get("id" if graph else "Id")
                if not identity:
                    raise OutlookCliError("Folder inventory returned no identity.")
                if identity in seen:
                    continue
                seen.add(identity)
                found.append({"id": identity, "displayName": item.get("displayName" if graph else "DisplayName", identity),
                              "totalItemCount": item.get("totalItemCount" if graph else "TotalItemCount")})
                if len(seen) > 10000:
                    raise OutlookCliError("Folder inventory safety limit reached.")
                if item.get("childFolderCount" if graph else "ChildFolderCount", 0):
                    queue.append(f"{root}/{quote(identity, safe='')}/childFolders")
            link = data.get("@odata.nextLink") or data.get("odata.nextLink")
            if not link:
                break
            link = validated_next_link(link, reader.base_url + "/", reader.base_url)
            if link in links:
                raise OutlookCliError("Repeated folder inventory page; store was not reconciled.")
            links.add(link)
            path, params = link, None
        else:
            raise OutlookCliError("Folder inventory page budget reached.")
    return found


@click.group()
def local():
    """Sync once; search, read, follow threads and inspect attachments offline."""


@local.command("sync")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest", show_default=True)
@click.option("--folder", "folders", multiple=True, help="Exact folder name or ID. Default: every discovered primary-mailbox folder.")
@click.option("--full", is_flag=True, help="Start fresh snapshots; keep existing data until each folder completes.")
@click.option("--page-size", type=click.IntRange(1, 200), default=100, show_default=True)
@click.option("--interval", type=click.FloatRange(min=0.25), default=0.25, help="Minimum seconds between requests (serial).")
@click.option("--max-pages", type=click.IntRange(1, 100000), default=10000)
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def sync(backend, folders, full, page_size, interval, max_pages, as_json, account_name):
    """Resume full copies, then retrieve only changed messages with folder delta."""
    profile = account.resolve_account_name(account_name)
    maybe_dry_run("local-sync", {"account": profile, "backend": backend, "folders": list(folders), "full": full})
    paths = account.get_account_paths(profile)
    started = time.monotonic()
    with file_lock(paths.cache_dir / "mail-sync.lock", timeout=30):
        store = MailStore(store_path(profile))
        graph = None
        try:
            cooldown = store.cooldown_remaining()
            if cooldown > 0:
                raise RateLimitError(f"Stored server cooldown active for {cooldown:.0f} seconds; no request was sent.", retry_after=cooldown)
            if backend == "graph":
                graph = GraphReader.for_account(profile)
                reader = GraphSyncReader(graph, interval=interval, page_size=page_size)
            else:
                reader = RestReader(_get_client(profile), interval=interval, page_size=page_size)
            try:
                available = discover_folders(reader)
            except RateLimitError as exc:
                store.set_cooldown(exc.retry_after or 60)
                raise
            except httpx.HTTPStatusError as exc:
                raise httpx.HTTPStatusError(
                    f"Folder discovery failed (HTTP {exc.response.status_code}); no folder was reconciled.",
                    request=exc.request, response=exc.response) from None
            except httpx.RequestError as exc:
                raise httpx.RequestError("Folder discovery connection failed; no folder was reconciled.", request=exc.request) from None
            selected = []
            for requested in folders:
                matches = [f for f in available if requested == f["id"] or requested.casefold() == f["displayName"].casefold()]
                if len(matches) != 1:
                    raise click.UsageError("Folder is missing or ambiguous; use its exact ID from folders.")
                if matches[0] not in selected:
                    selected.append(matches[0])
            if not folders:
                selected = available
            store.reconcile_folders(backend, {f["id"] for f in available})
            outcomes = []
            for position, folder in enumerate(selected, 1):
                try:
                    def progress(pages, records, position=position):
                        if pages == 1 or pages % 10 == 0:
                            click.echo(f"Sync folder {position}/{len(selected)}: {pages} pages, {records} records", err=True)
                    outcome = sync_folder(store, reader, folder, full=full, max_pages=max_pages, on_page=progress)
                    outcomes.append({"folder_id": folder["id"], "name": folder["displayName"], **outcome})
                except Exception as exc:  # noqa: BLE001 - return safe partial diagnostics and stop all requests
                    code = error_code_for_exception(exc)
                    # Provider exceptions can contain opaque cursor URLs. Keep diagnostics content-free.
                    status = exc.response.status_code if isinstance(exc, httpx.HTTPStatusError) else None
                    message = f"Sync stopped ({code}" + (f", HTTP {status}" if status else "") + "); rerun to resume."
                    store.failed(backend, folder["id"], folder["displayName"], code)
                    if isinstance(exc, RateLimitError):
                        store.set_cooldown(exc.retry_after or 60)
                    outcomes.append({"folder_id": folder["id"], "ok": False, "error": {"code": code, "message": message}})
                    break  # Never keep hitting a server after throttling/auth/transport failure.
            failed = any(not item["ok"] for item in outcomes)
            meta = {**store.status(backend), "elapsed_seconds": round(time.monotonic() - started, 3),
                    "requests": reader.requests, "response_bytes": reader.response_bytes,
                    "selected_folders": len(selected), "completed_folders": sum(item["ok"] for item in outcomes),
                    "remaining_folders": len(selected) - sum(item["ok"] for item in outcomes), "partial": failed}
            error = {"code": "partial_failure", "message": "Sync stopped; previous snapshots and committed pages retained. Rerun to resume."} if failed else None
            click.echo(to_json_envelope(Page(outcomes, meta=meta), ok=not failed, error=error))
            if failed:
                raise click.exceptions.Exit(1)
        finally:
            store.close()
            if graph:
                graph.close()


@local.command("search")
@click.argument("query", default="")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest")
@click.option("--max", "--limit", "-n", "limit", default=25, type=click.IntRange(1, 100000))
@click.option("--folder")
@click.option("--from", "sender", help="Exact sender email address")
@click.option("--to", "recipient", help="Exact recipient (To/CC/BCC)")
@click.option("--person", help="Sender/recipient email address or folded display-name fragment")
@click.option("--domain")
@click.option("--after", help="Inclusive UTC date/time")
@click.option("--before", help="Exclusive UTC date/time")
@click.option("--conversation", "--thread", "conversation")
@click.option("--has-attachments", is_flag=True)
@click.option("--match", "match_mode", type=click.Choice(["exact", "prefix", "stem"]), default="prefix", show_default=True)
@click.option("--require-complete", is_flag=True)
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def search(query, backend, limit, folder, sender, recipient, person, domain, after, before, conversation,
           has_attachments, match_mode, require_complete, as_json, account_name):
    """Offline BM25 search. Quotes mean phrases; field:value terms add filters."""
    store = _store(account_name)
    try:
        rows, meta = store.query(query, backend=backend, limit=limit, folder=folder, sender=sender, recipient=recipient,
                                 person=person, domain=domain, after=after, before=before, conversation=conversation,
                                 has_attachments=True if has_attachments else None, match_mode=match_mode, require_complete=require_complete)
        click.echo(to_json_envelope(Page(rows, meta=meta)))
    finally:
        store.close()


@local.command("read")
@click.argument("message_id")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest")
@click.option("--html", is_flag=True, help="Original HTML, only when retained by legacy import")
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def read(message_id, backend, html, as_json, account_name):
    """Read a full locally stored message as text. Never marks it read."""
    store = _store(account_name)
    try:
        click.echo(to_json_envelope(store.read(message_id, backend=backend, body_format="html" if html else "text")))
    finally:
        store.close()


@local.command("thread")
@click.argument("message_or_conversation_id")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest")
@click.option("--max", "--limit", "-n", "limit", default=100, type=click.IntRange(1, 100000))
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def thread(message_or_conversation_id, backend, limit, as_json, account_name):
    """Follow a message or conversation ID across locally stored folders."""
    store = _store(account_name)
    try:
        rows, meta = store.thread(message_or_conversation_id, backend=backend, limit=limit)
        click.echo(to_json_envelope(Page(rows, meta=meta)))
    finally:
        store.close()


@local.command("related")
@click.argument("message_id")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest")
@click.option("--max", "--limit", "-n", "limit", default=25, type=click.IntRange(1, 100000))
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def related(message_id, backend, limit, as_json, account_name):
    """Find other correspondence with the message sender, offline."""
    store = _store(account_name)
    try:
        rows, meta = store.related(message_id, backend=backend, limit=limit)
        click.echo(to_json_envelope(Page(rows, meta=meta)))
    finally:
        store.close()


@local.command("attachments")
@click.argument("message_id")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest")
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def attachments(message_id, backend, as_json, account_name):
    """List cached attachment metadata. Binary files are not fetched by sync."""
    store = _store(account_name)
    try:
        rows, meta = store.attachments(message_id, backend=backend, include_meta=True)
        click.echo(to_json_envelope(Page(rows, meta=meta)))
    finally:
        store.close()


@local.command("status")
@click.option("--backend", type=click.Choice(["rest", "graph"]), default="rest")
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def status(backend, as_json, account_name):
    """Inspect coverage, checkpoints, freshness and disk size without network."""
    store = _store(account_name)
    try:
        click.echo(to_json_envelope(store.status(backend)))
    finally:
        store.close()


@local.command("import-index")
@click.option("--source", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def import_index(source, as_json, account_name):
    """Import the old index beside it. Never delete or alter the source."""
    profile = account.resolve_account_name(account_name)
    source = source or account.get_account_paths(profile).cache_dir / "index.sqlite3"
    maybe_dry_run("local-import-index", {"source": str(source), "destination": str(store_path(profile))})
    with file_lock(store_path(profile).parent / "mail-sync.lock", timeout=30):
        store = MailStore(store_path(profile))
        try:
            click.echo(to_json_envelope(store.migrate_legacy(source)))
        finally:
            store.close()


@local.command("compact")
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def compact(as_json, account_name):
    """Reclaim superseded bodies and SQLite free pages in the new store."""
    profile = account.resolve_account_name(account_name)
    maybe_dry_run("local-compact", {"path": str(store_path(profile))})
    with file_lock(store_path(profile).parent / "mail-sync.lock", timeout=30):
        store = MailStore(store_path(profile))
        try:
            click.echo(to_json_envelope(store.compact()))
        finally:
            store.close()


@local.command("purge")
@click.option("-y", "--yes", is_flag=True)
@click.option("--include-legacy-index", is_flag=True, help="Explicitly remove index.sqlite3 and its sidecars too")
@click.option("--json", "as_json", is_flag=True)
@account_option
@_handle_api_error
def purge(yes, include_legacy_index, as_json, account_name):
    """Remove the new local mail store and sidecars. Legacy index is preserved."""
    profile = account.resolve_account_name(account_name)
    path = store_path(profile)
    legacy = path.parent / "index.sqlite3"
    maybe_dry_run("local-purge", {"path": str(path), "legacy_index_retained": not include_legacy_index,
                                 "legacy_path": str(legacy) if include_legacy_index else None})
    confirm_action("Delete local mail, including the legacy index?" if include_legacy_index else "Delete this local mail store?",
                   yes=yes, action="purge local mail")
    with ExitStack() as locks:
        locks.enter_context(file_lock(path.parent / "mail-sync.lock", timeout=30))
        if include_legacy_index:
            locks.enter_context(file_lock(path.parent / "index-sync.lock", timeout=30))
        targets = [Path(str(base) + suffix) for base in ([path, legacy] if include_legacy_index else [path])
                   for suffix in ("", "-wal", "-shm")]
        if any(target.is_symlink() or (target.exists() and not target.is_file()) for target in targets):
            raise OutlookCliError("Refusing to purge symlinks or non-regular mail-store files.")
        removed = []
        for target in targets:
            if target.exists():
                target.unlink()
                removed.append(str(target))
        click.echo(to_json_envelope({"removed": removed, "legacy_index_retained": not include_legacy_index}))
