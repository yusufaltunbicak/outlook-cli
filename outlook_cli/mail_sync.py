"""Read-only, serial, resumable mailbox synchronization.

REST uses the legacy track-changes protocol, including its initial seed round.
Only complete pages and their continuation are committed; opaque cursors never
appear in CLI output. Graph remains a separately authenticated alternative.
"""
from __future__ import annotations

import re
import time
from dataclasses import asdict
from urllib.parse import quote

import httpx

from .constants import BASE_URL
from .exceptions import OutlookCliError, RateLimitError, ResourceNotFoundError
from .graph import FIELDS as GRAPH_FIELDS, GRAPH_URL, graph_to_record
from .models import Email
from .pagination import validated_next_link
from .transport import request_response

REST_FIELDS = (
    "Id,Subject,From,ToRecipients,CcRecipients,BccRecipients,ReplyTo,"
    "ReceivedDateTime,SentDateTime,BodyPreview,Body,IsRead,HasAttachments,"
    "ConversationId,Categories,ParentFolderId,LastModifiedDateTime,InternetMessageId"
)
ATTACHMENT_FIELDS = "Id,Name,Size,ContentType,IsInline"
GRAPH_EXTRA = ",body,bccRecipients,replyTo,sentDateTime,internetMessageId"


def _addresses(values):
    return [{"name": x.get("EmailAddress", {}).get("Name", ""),
             "address": x.get("EmailAddress", {}).get("Address", "")} for x in values or []]


def rest_record(value):
    value = dict(value)
    # Drafts/system items may legitimately carry null sender/body fields.
    for field in ("From", "Body"):
        value[field] = value.get(field) or {}
    for field in ("ToRecipients", "CcRecipients", "Categories"):
        value[field] = value.get(field) or []
    record = asdict(Email.from_api(value))
    record.update(bcc=_addresses(value.get("BccRecipients")), reply_to=_addresses(value.get("ReplyTo")),
                  sent=value.get("SentDateTime"), modified=value.get("LastModifiedDateTime"),
                  internet_message_id=value.get("InternetMessageId", ""),
                  etag=value.get("@odata.etag", ""), attachments=[
                      {"id": a.get("Id"), "name": a.get("Name", ""), "size": a.get("Size", 0),
                       "content_type": a.get("ContentType", ""), "is_inline": a.get("IsInline", False),
                       "type": a.get("@odata.type", "")} for a in value.get("Attachments", [])])
    return record


def removed_id(value):
    """The legacy REST tombstone is lowercase id=Messages('...')."""
    if "@removed" not in value and value.get("reason") not in {"deleted", "changed", "moved"}:
        return None
    identity = value.get("Id") or value.get("id")
    if not isinstance(identity, str):
        raise OutlookCliError("A removed message has no identity; checkpoint retained.")
    match = re.fullmatch(r"Messages\('((?:[^']|'')*)'\)", identity, re.IGNORECASE)
    return match.group(1).replace("''", "'") if match else identity


class RestReader:
    backend = "rest"
    base_url = BASE_URL
    fields = REST_FIELDS

    def __init__(self, client, *, interval=0.25, page_size=100):
        self.client = client
        self.interval = interval
        self.page_size = page_size
        self.last_request = 0.0
        self.requests = 0
        self.response_bytes = 0

    def get(self, path, params=None):
        if path.startswith(("https:", "http:", "//")):
            path = validated_next_link(path, self.base_url, self.base_url)
        remaining = self.interval - (time.monotonic() - self.last_request)
        if remaining > 0:
            time.sleep(remaining)
        self.last_request = time.monotonic()
        response = request_response(
            self.client._client, "GET", path, params=params,
            headers={"Prefer": f'odata.track-changes, odata.maxpagesize={self.page_size}, outlook.body-content-type="text"'},
            refresh=self.client._refresh_token if self.client._refresh else None,
            account_name=self.client.account_name,
        )
        self.requests += 1
        self.response_bytes += len(response.content)
        data = response.json()
        data["_tracking"] = "odata.track-changes" in response.headers.get("Preference-Applied", "").lower()
        return data

    def initial(self, folder_id):
        return f"/MailFolders/{quote(folder_id, safe='')}/messages", {
            "$select": self.fields, "$top": self.page_size,
            "$expand": f"Attachments($select={ATTACHMENT_FIELDS})",
        }

    def hydrate(self, identity):
        return self.get(f"/messages/{quote(identity, safe='')}", {
            "$select": self.fields, "$expand": f"Attachments($select={ATTACHMENT_FIELDS})"})

    def record(self, message):
        return rest_record(message)


class GraphSyncReader:
    backend = "graph"
    base_url = GRAPH_URL
    fields = GRAPH_FIELDS + GRAPH_EXTRA

    def __init__(self, reader, *, interval=0.25, page_size=100):
        self.reader = reader
        self.interval, self.page_size = interval, page_size
        self.last_request = 0.0
        self.requests = 0
        self.response_bytes = None  # GraphReader's JSON adapter does not expose wire bytes.

    def get(self, path, params=None):
        remaining = self.interval - (time.monotonic() - self.last_request)
        if remaining > 0:
            time.sleep(remaining)
        self.last_request = time.monotonic()
        result = self.reader.get(path, params=params)
        self.requests += 1
        return result

    def initial(self, folder_id):
        return f"/me/mailFolders/{quote(folder_id, safe='')}/messages/delta", {
            "$select": self.fields, "$top": self.page_size}

    def hydrate(self, identity):
        return self.get(f"/me/messages/{quote(identity, safe='')}", {
            "$select": self.fields,
            "$expand": "attachments($select=id,name,size,contentType,isInline)"})

    def record(self, message):
        result = graph_to_record(message)
        def addresses(items):
            return [{"name": x.get("emailAddress", {}).get("name", ""),
                     "address": x.get("emailAddress", {}).get("address", "")} for x in items or []]
        result.update(bcc=addresses(message.get("bccRecipients")), reply_to=addresses(message.get("replyTo")),
                      sent=message.get("sentDateTime"), modified=message.get("lastModifiedDateTime"),
                      internet_message_id=message.get("internetMessageId", ""), attachments=[
                          {"id": a.get("id"), "name": a.get("name", ""), "size": a.get("size", 0),
                           "content_type": a.get("contentType", ""), "is_inline": a.get("isInline", False),
                           "type": a.get("@odata.type", "")} for a in message.get("attachments", [])])
        return result


def sync_folder(store, reader, folder, *, full=False, max_pages=10000, on_page=None):
    """Commit one page at a time and resume an interrupted initial/delta round."""
    backend = reader.backend
    identity, name = folder["id"], folder["displayName"]
    signature = f"text+attachments:v1:{backend}:{reader.page_size}"
    previous = store.folder_state(backend, identity)
    pending = bool(previous.get("pending_url")) and previous.get("options_signature") == signature
    rebuild = full or (bool(previous.get("pending_full")) if pending else not previous.get("cursor")) or previous.get("options_signature") != signature
    state = store.begin_sync(backend, identity, name, full=rebuild, include_body=True,
                             options_signature=signature, restart=full)
    initial_path, initial_params = reader.initial(identity)
    path = state.get("pending_url") or (None if rebuild else state.get("cursor")) or initial_path
    params = initial_params if path == initial_path else None
    seed_done = bool(state.get("seed_done")) or not rebuild or backend == "graph"
    fallback_snapshot = bool(state.get("snapshot_mode"))
    pages = records = removed = 0
    seen_links = set()
    started = time.monotonic()
    reset_attempted = False
    while pages < max_pages:
        try:
            response = reader.get(path, params=params)
        except httpx.HTTPStatusError as exc:
            # Invalidated server state must not destroy the old complete snapshot.
            try:
                provider_code = exc.response.json().get("error", {}).get("code", "")
            except (ValueError, AttributeError):
                provider_code = ""
            expired = exc.response.status_code == 410 or (
                exc.response.status_code in {400, 404} and provider_code.lower() in {
                    "syncstatenotfound", "errorinvalidsyncstatedata", "invaliddeltatoken", "resyncrequired"})
            if expired and not reset_attempted:
                state = store.begin_sync(backend, identity, name, full=True, include_body=True,
                                         options_signature=signature, restart=True)
                path, params = initial_path, initial_params
                rebuild, seed_done, reset_attempted = True, backend == "graph", True
                seen_links.clear()
                fallback_snapshot = False
                continue
            raise
        batch = response.get("value")
        if not isinstance(batch, list):
            raise OutlookCliError("Invalid sync page; checkpoint retained.")
        next_link = response.get("@odata.nextLink") or response.get("odata.nextLink")
        delta_link = response.get("@odata.deltaLink") or response.get("odata.deltaLink")
        for link in (next_link, delta_link):
            if link:
                validated_next_link(link, reader.base_url + "/", reader.base_url)
        if path == initial_path and params is not None and not seed_done and not response.get("_tracking"):
            fallback_snapshot = True
        initial_seed = backend == "rest" and rebuild and not seed_done and not fallback_snapshot
        continuation = next_link or (delta_link if initial_seed else None)
        if initial_seed and not continuation:
            raise OutlookCliError("REST tracking returned no seed checkpoint; retry with --full.")
        if not continuation and not delta_link and not fallback_snapshot:
            raise OutlookCliError("Sync page returned no delta checkpoint; previous checkpoint retained.")
        if continuation in seen_links:
            raise OutlookCliError("Repeated sync continuation; checkpoint retained.")
        current, deleted = [], []
        identities = set()
        for message in batch:
            hydrated = False
            tombstone = removed_id(message)
            if tombstone:
                # Hydrate delta tombstones too: a later update may appear before an older delete.
                if not rebuild:
                    try:
                        message = reader.hydrate(tombstone)
                        hydrated = True
                    except (ResourceNotFoundError, httpx.HTTPStatusError) as exc:
                        if isinstance(exc, httpx.HTTPStatusError) and exc.response.status_code != 404:
                            raise
                        deleted.append(tombstone)
                        continue
                else:
                    deleted.append(tombstone)
                    continue
            msg_id = message.get("Id") or message.get("id")
            if not msg_id:
                raise OutlookCliError("Sync message has no identity; checkpoint retained.")
            # Delta ordering is not authoritative. Resolve each change against current state.
            required = {"Id", "Subject", "Body", "From", "ToRecipients", "ParentFolderId"} if backend == "rest" else {"id", "subject", "body", "from", "toRecipients", "parentFolderId"}
            if not hydrated and (not rebuild or not required.issubset(message) or msg_id in identities or backend == "graph"):
                try:
                    message = reader.hydrate(msg_id)
                except (ResourceNotFoundError, httpx.HTTPStatusError) as exc:
                    if isinstance(exc, httpx.HTTPStatusError) and exc.response.status_code != 404:
                        raise
                    deleted.append(msg_id)
                    continue
            if not required.issubset(message):
                raise OutlookCliError("Hydrated message is incomplete; checkpoint retained.")
            parent = message.get("ParentFolderId") or message.get("parentFolderId")
            if parent and parent != identity:
                deleted.append(msg_id)
                continue
            identities.add(msg_id)
            current.append(reader.record(message))
        complete = continuation is None
        next_seed_done = seed_done or bool(initial_seed and delta_link)
        store.apply_page(backend, identity, name, current, removed=deleted,
                         next_url=continuation, cursor=delta_link if complete else None,
                         complete=complete, seed_done=next_seed_done, snapshot_mode=fallback_snapshot)
        pages += 1
        records += len(current)
        removed += len(deleted)
        if on_page:
            on_page(pages, records)
        if complete:
            return {"ok": True, "records": records, "removed": removed, "pages": pages,
                    "mode": "snapshot" if fallback_snapshot else ("full" if rebuild else "delta"),
                    "resumed": bool(previous.get("pending_url")),
                    "elapsed_seconds": round(time.monotonic() - started, 3)}
        seen_links.add(continuation)
        path, params, seed_done = continuation, None, next_seed_done
    raise OutlookCliError("Sync page budget reached; rerun the same command to resume.")
