"""Explicit local indexing; read operations never silently re-sync the mailbox."""
from __future__ import annotations

from dataclasses import asdict
from urllib.parse import quote

import click
import httpx

from .. import account
from ..exceptions import OutlookCliError, error_code_for_exception
from ..graph import GraphReader, graph_login as login_graph, graph_to_record, materialize_changes
from ..index_store import MailIndex
from ..locking import file_lock
from ..models import Email
from ..pagination import Page, paginate
from ..serialization import to_json_envelope
from ._common import _get_client, _handle_api_error, account_option, is_no_input_mode, maybe_dry_run

REST_FIELDS = "Id,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,BodyPreview,IsRead,HasAttachments,ConversationId,Categories"


def _store(profile: str):
    return MailIndex(account.get_account_paths(profile).cache_dir / "index.sqlite3")


def _rest_folders(client):
    queue,result,seen=["/MailFolders"],[],set()
    while queue:
        if len(seen)>10000:
            raise OutlookCliError("Folder hierarchy exceeded its safety limit.")
        page=paginate(client._get,queue.pop(0),{"$top":100},limit=None)
        if not page.meta["complete"]:
            raise OutlookCliError("Incomplete folder hierarchy; index was not changed.")
        for folder in page:
            if folder["Id"] in seen:
                continue
            seen.add(folder["Id"])
            result.append({"id":folder["Id"],"displayName":folder.get("DisplayName",folder["Id"])})
            if folder.get("ChildFolderCount",0):
                queue.append(f"/MailFolders/{quote(folder['Id'],safe='')}/childfolders")
    return result


@click.group()
def index():
    """Synchronize and search an account-local mail index."""


@index.command("sync")
@click.option("--backend",type=click.Choice(["rest","graph"]),default="rest",show_default=True)
@click.option("--folder", "folders", multiple=True, help="Exact folder name or ID; repeatable. Default: all folders and children.")
@click.option("--include-body",is_flag=True,help="Store bodies too; default metadata/preview only")
@click.option("--full",is_flag=True,help="Reset Graph delta and rebuild selected folders")
@click.option("--json","as_json",is_flag=True)
@account_option
@_handle_api_error
def sync(backend,folders,include_body,full,as_json,account_name):
    """Refresh local data; REST scans metadata, Graph subsequently uses delta."""
    profile=account.resolve_account_name(account_name)
    maybe_dry_run("index-sync",{"account":profile,"backend":backend,"folders":list(folders),"include_body":include_body})
    paths=account.get_account_paths(profile)
    outcomes=[]
    graph=None
    with file_lock(paths.cache_dir / "index-sync.lock",timeout=30):
        store=_store(profile)
        try:
            if backend=="graph":
                graph=GraphReader.for_account(profile)
                available=graph.folders()
            else:
                client=_get_client(profile)
                available=_rest_folders(client)
            selected=[]
            if folders:
                for requested in folders:
                    matches=[f for f in available if requested==f["id"] or requested.casefold()==f["displayName"].casefold()]
                    if len(matches)!=1:
                        raise click.UsageError(f"Folder '{requested}' is missing or ambiguous; use its exact ID.")
                    if matches[0] not in selected:
                        selected.append(matches[0])
            else:
                selected=available
            for folder in selected:
                identity,name=folder["id"],folder["displayName"]
                try:
                    state=store.folder_state(backend,identity)
                    if graph:
                        cursor=state.get("cursor") if not full and bool(state.get("include_body"))==include_body else None
                        try:
                            changes,checkpoint=graph.delta(identity,cursor,include_body=include_body)
                        except httpx.HTTPStatusError as exc:
                            if cursor and exc.response.status_code==410:
                                changes,checkpoint=graph.delta(identity,include_body=include_body)
                                cursor=None
                            else:
                                raise
                        # Repeated/partial entries have no guaranteed order: read their current state.
                        changes=materialize_changes(graph,changes,identity,include_body=include_body)
                        records=[graph_to_record(m) for m in changes if "@removed" not in m]
                        removed=[m["id"] for m in changes if "@removed" in m]
                        store.apply(backend,identity,name,records,full=cursor is None,cursor=checkpoint,removed=removed,include_body=include_body)
                        mode="delta" if cursor else "full"
                    else:
                        page=paginate(client._get,f"/MailFolders/{quote(identity,safe='')}/messages",
                            {"$top":100,"$select":REST_FIELDS+(",Body" if include_body else "")},limit=None,max_pages=1000)
                        if not page.meta["complete"]:
                            raise OutlookCliError(f"Incomplete folder snapshot: {page.meta['truncated_reason']}")
                        records=[asdict(Email.from_api(m)) for m in page]
                        store.apply(backend,identity,name,records,full=True,include_body=include_body)
                        mode="full_metadata" if not include_body else "full"
                    outcomes.append({"folder_id":identity,"name":name,"ok":True,"records":len(records),"mode":mode})
                except Exception as exc:
                    code=error_code_for_exception(exc)
                    store.failed(backend,identity,name,code)
                    outcomes.append({"folder_id":identity,"name":name,"ok":False,"error":{"code":code,"message":str(exc)}})
            if not folders:
                store.reconcile_folders(backend,{f["id"] for f in available})
            meta=store.status(backend)
        finally:
            store.close()
            if graph:
                graph.close()
    failed=any(not r["ok"] for r in outcomes)
    page=Page(outcomes,meta={**meta,"partial":failed})
    # Explicit partial result must not look like a successful full snapshot.
    error={"code":"partial_failure","message":"Some folders did not synchronize; previous data retained."} if failed else None
    click.echo(to_json_envelope(page,ok=not failed,error=error))
    if failed:
        raise click.exceptions.Exit(1)


@index.command("search")
@click.argument("query",default="")
@click.option("--backend",type=click.Choice(["rest","graph"]),default="rest")
@click.option("--max","--limit","-n","limit",default=25,type=click.IntRange(1,100000))
@click.option("--folder")
@click.option("--from","sender",help="Exact sender address")
@click.option("--to","recipient",help="Exact To/CC address")
@click.option("--domain",help="Exact sender or recipient domain")
@click.option("--after",help="Inclusive UTC date/time")
@click.option("--before",help="Exclusive UTC date/time")
@click.option("--conversation",help="Exact conversation ID")
@click.option("--require-complete",is_flag=True)
@click.option("--json","as_json",is_flag=True)
@account_option
@_handle_api_error
def search(query,backend,limit,folder,sender,recipient,domain,after,before,conversation,require_complete,as_json,account_name):
    """Search local data without any network or auth. Quotes select phrases."""
    store=_store(account.resolve_account_name(account_name))
    try:
        rows,meta=store.query(query,backend=backend,limit=limit,folder=folder,sender=sender,recipient=recipient,
            domain=domain,after=after,before=before,conversation=conversation,require_complete=require_complete)
        click.echo(to_json_envelope(Page(rows,meta=meta)))
    finally:
        store.close()


@index.command("status")
@click.option("--backend",type=click.Choice(["rest","graph"]),default="rest")
@click.option("--json","as_json",is_flag=True)
@account_option
@_handle_api_error
def status(backend,as_json,account_name):
    """Show indexed scope and freshness without contacting Outlook."""
    store=_store(account.resolve_account_name(account_name))
    try:
        click.echo(to_json_envelope(store.status(backend)))
    finally:
        store.close()


@click.command("graph-login")
@click.option("--client-id",required=True,envvar="OUTLOOK_GRAPH_CLIENT_ID")
@click.option("--tenant",default="organizations",envvar="OUTLOOK_GRAPH_TENANT")
@click.option("--timeout",default=300,type=click.IntRange(15,900))
@click.option("--json","as_json",is_flag=True)
@account_option
@_handle_api_error
def graph_login(client_id,tenant,timeout,as_json,account_name):
    """Explicit optional Graph device login; requires an Entra public client app."""
    maybe_dry_run("graph-login",{"tenant":tenant,"client_id":client_id})
    if is_no_input_mode():
        from ..exceptions import AuthRequiredError
        raise AuthRequiredError("graph-login is interactive and cannot run with --no-input.")
    result=login_graph(account.resolve_account_name(account_name),client_id,tenant,timeout)
    click.echo(to_json_envelope(result))
