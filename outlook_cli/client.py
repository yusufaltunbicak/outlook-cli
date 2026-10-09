from __future__ import annotations

import json
import threading
import uuid
from pathlib import Path

import httpx

from urllib.parse import quote

from . import account as account_service
from .constants import ATTACHMENT_SIZE_THRESHOLD, BASE_URL, DEFERRED_SEND_PROPERTY_ID, OWA_SERVICE_URL, USER_AGENT
from .exceptions import AccountError, OutlookCliError, ResourceNotFoundError, TokenExpiredError
from .models import Attachment, Contact, Email, Event, Folder
from .pagination import Page, paginate
from .state import IDStore, atomic_json_write, locked_file
from .transport import request_json, request_response

import html as _html_mod


def _plain_text_to_html(text: str) -> str:
    """Convert plain text to basic HTML, preserving line breaks.

    Escapes HTML special characters and replaces newlines with <br> tags.
    """
    escaped = _html_mod.escape(text)
    return escaped.replace("\n", "<br>\n")


def _build_query_params(
    unread_only: bool = False,
    filter_from: str | None = None,
    filter_subject: str | None = None,
    filter_after: str | None = None,
    filter_before: str | None = None,
    filter_has_attachments: bool = False,
    filter_category: str | None = None,
) -> tuple[str, str, bool]:
    """Build $filter and $search params.

    REST v2 limitations:
    - $filter and $search can't be combined
    - $filter doesn't support contains() on From
    - $search KQL supports from:, subject:, hasattachments:, received:

    Strategy: if text filters (from/subject) are used, build a KQL $search.
    Otherwise use $filter for IsRead/date (which supports $orderby).

    Returns (filter_str, search_str, needs_search).
    When needs_search is True, $orderby must be omitted.
    """
    has_text_filters = any([filter_from, filter_subject, filter_has_attachments])

    if has_text_filters:
        # Use $search with KQL — can't combine with $filter
        kql_parts: list[str] = []
        if filter_from:
            kql_parts.append(f"from:{filter_from}")
        if filter_subject:
            kql_parts.append(f"subject:{filter_subject}")
        if filter_has_attachments:
            kql_parts.append("hasattachments:true")
        if unread_only:
            kql_parts.append("isread:false")
        if filter_after:
            kql_parts.append(f"received>={filter_after}")
        if filter_before:
            kql_parts.append(f"received<={filter_before}")
        if filter_category:
            kql_parts.append(f'category:"{filter_category}"')
        return "", f'"{" ".join(kql_parts)}"', True

    # Pure $filter — supports $orderby
    filter_parts: list[str] = []
    if unread_only:
        filter_parts.append("IsRead eq false")
    if filter_after:
        filter_parts.append(f"ReceivedDateTime ge {filter_after}T00:00:00Z")
    if filter_before:
        filter_parts.append(f"ReceivedDateTime lt {filter_before}T23:59:59Z")
    if filter_category:
        escaped_category = filter_category.replace("'", "''")
        filter_parts.append(f"Categories/any(c:c eq '{escaped_category}')")
    return " and ".join(filter_parts), "", False


class OutlookClient:
    """HTTP client for Outlook REST API v2."""

    MAX_ID_MAP_SIZE = None  # Compatibility only: display references never expire.

    def __init__(self, token: str, account_name: str | None = None, refresh=None):
        self.account_name = account_service.resolve_account_name(account_name)
        self._paths = account_service.get_account_paths(self.account_name)
        self._token = token
        self._refresh = refresh
        self._folder_cache = {}
        self._folder_lock = threading.RLock()
        self._pool_lock = threading.Lock()
        self._owa_client = None
        self._upload_client = None
        self._id_store = None
        self._client = httpx.Client(
            base_url=BASE_URL,
            headers={
                "Authorization": f"Bearer {token}",
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json",
            },
            timeout=30,
        )
        # Resolve durable references on demand: startup must not load an ever-growing map.
        self._id_map: dict[str, str] = {}

    # ------------------------------------------------------------------
    # Mail
    # ------------------------------------------------------------------

    def get_messages(
        self, folder: str = "Inbox", top: int = 25, skip: int = 0,
        unread_only: bool = False, filter_from: str | None = None,
        filter_subject: str | None = None, filter_after: str | None = None,
        filter_before: str | None = None, filter_has_attachments: bool = False,
        filter_category: str | None = None, filter_no_category: bool = False,
        select: str | None = None, all_pages: bool = False,
    ) -> Page:
        folder_id = self._resolve_folder(folder)
        filter_str, search_str, needs_search = _build_query_params(
            unread_only, filter_from, filter_subject, filter_after,
            filter_before, filter_has_attachments, filter_category,
        )
        params = {"$top": min(top * 3 if filter_no_category else top, 1000)}
        if needs_search:
            params["$search"] = search_str
        else:
            params.update({"$skip": skip, "$orderby": "ReceivedDateTime desc"})
            if filter_str:
                params["$filter"] = filter_str
        if select:
            fields = list(dict.fromkeys(["Id", *select.split(","), *(["Categories"] if filter_no_category else [])]))
            params["$select"] = ",".join(fields)
        raw = paginate(self._get, f"/MailFolders/{folder_id}/messages", params,
                       limit=None if all_pages else top, search=needs_search,
                       predicate=(lambda item: not item.get("Categories")) if filter_no_category else None)
        messages = Page((Email.from_api(item) for item in raw), meta=raw.meta)
        self._assign_display_nums(messages)
        return messages

    def get_message(self, message_id: str) -> Email:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}")
        email = Email.from_api(resp)
        email.display_num = int(message_id) if message_id.isdigit() else 0
        return email

    def get_thread(self, message_id: str, max_messages: int = 50) -> Page:
        """Search the conversation, keeping its seed and reporting search coverage."""
        import re
        email = self.get_message(message_id)
        base_subject = re.sub(
            r"^(?:(?:Re|Fwd|İlt|Ynt|Fw|AW|SV|VS)\s*:\s*)+", "",
            email.subject, flags=re.IGNORECASE,
        ).strip()
        if not email.conversation_id or not base_subject:
            return Page([email], meta={"complete": False, "has_more": None,
                        "truncated_reason": "conversation_identity_unavailable", "pages": 0, "returned_count": 1})
        raw = paginate(self._get, "/messages",
                       {"$search": f'"subject:{base_subject}"', "$top": min(max_messages, 1000)},
                       limit=max_messages, search=True,
                       predicate=lambda m: m.get("ConversationId") == email.conversation_id)
        by_id = {item["Id"]: Email.from_api(item) for item in raw}
        by_id[email.id] = email
        items = list(by_id.values())
        if len(items) > max_messages:
            items = [m for m in items if m.id != email.id][:max(0, max_messages - 1)] + [email]
            raw.meta.update(complete=False, has_more=True, truncated_reason="limit")
        from datetime import timezone
        items.sort(key=lambda m: m.received.replace(tzinfo=timezone.utc) if m.received.tzinfo is None else m.received)
        thread = Page(items, meta=raw.meta)
        thread.meta["returned_count"] = len(thread)
        thread.meta["scope"] = "subject_search_filtered_by_conversation"
        self._assign_display_nums(thread)
        return thread

    def send_mail(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
        send_at: str | None = None,
    ) -> None:
        content = body if html else _plain_text_to_html(body)
        message: dict = {
            "Subject": subject,
            "Body": {
                "ContentType": "HTML",
                "Content": content,
            },
            "ToRecipients": [
                {"EmailAddress": {"Address": addr}} for addr in to
            ],
        }
        if cc:
            message["CcRecipients"] = [
                {"EmailAddress": {"Address": addr}} for addr in cc
            ]
        if send_at:
            message["SingleValueExtendedProperties"] = [{
                "PropertyId": DEFERRED_SEND_PROPERTY_ID,
                "Value": send_at,
            }]
        self._post("/sendmail", json={"Message": message})

    def create_draft(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
    ) -> Email:
        content = body if html else _plain_text_to_html(body)
        payload: dict = {
            "Subject": subject,
            "Body": {
                "ContentType": "HTML",
                "Content": content,
            },
            "ToRecipients": [
                {"EmailAddress": {"Address": addr}} for addr in to
            ],
        }
        if cc:
            payload["CcRecipients"] = [
                {"EmailAddress": {"Address": addr}} for addr in cc
            ]
        data = self._post("/messages", json=payload)
        return Email.from_api(data)

    def send_draft(self, message_id: str) -> None:
        real_id = self._resolve_id(message_id)
        self._post(f"/messages/{real_id}/send")

    def reply(self, message_id: str, comment: str, reply_all: bool = False) -> None:
        # Comment field only supports plain text; use draft flow to preserve
        # line breaks by converting to HTML.
        draft = self.create_reply_draft(message_id, comment=comment, reply_all=reply_all)
        self.send_draft(draft.id)

    def create_reply_draft(
        self,
        message_id: str,
        comment: str = "",
        reply_all: bool = False,
        html: bool = False,
    ) -> Email:
        real_id = self._resolve_id(message_id)
        action = "createreplyall" if reply_all else "createreply"

        if comment:
            # createReply's Comment field only supports plain text and drops
            # line breaks.  Always create an empty draft first, then prepend
            # the user's body (converted to HTML when needed) before the
            # quoted original message.
            data = self._post(f"/messages/{real_id}/{action}", json={})
            draft_id = data["Id"]
            original_body = data.get("Body", {}).get("Content", "")
            body_html = comment if html else _plain_text_to_html(comment)
            if "<body>" in original_body:
                combined = original_body.replace("<body>", f"<body>{body_html} ", 1)
            else:
                combined = body_html + original_body
            data = self._patch(f"/messages/{draft_id}", json={
                "Body": {"ContentType": "HTML", "Content": combined},
            })
        else:
            data = self._post(f"/messages/{real_id}/{action}", json={})

        return Email.from_api(data)

    def forward(self, message_id: str, to: list[str], comment: str = "") -> None:
        # Comment field only supports plain text; use draft flow to preserve
        # line breaks by converting to HTML.
        draft = self.create_forward_draft(message_id, to, comment=comment)
        self.send_draft(draft.id)

    def move_message(self, message_id: str, destination_folder: str) -> Email:
        real_id = self._resolve_id(message_id)
        folder_id = self._resolve_folder(destination_folder)
        resp = self._post(
            f"/messages/{real_id}/move",
            json={"DestinationId": folder_id},
        )
        email = Email.from_api(resp)
        if email.id:
            email.display_num = self._store().replace(real_id, email.id)
            self._id_map = {number: email.id if value == real_id else value for number, value in self._id_map.items()}
            self._id_map[str(email.display_num)] = email.id
        return email

    def copy_message(self, message_id: str, destination_folder: str) -> Email:
        real_id = self._resolve_id(message_id)
        folder_id = self._resolve_folder(destination_folder)
        resp = self._post(
            f"/messages/{real_id}/copy",
            json={"DestinationId": folder_id},
        )
        return Email.from_api(resp)

    def _resolve_folder(self, name_or_id: str) -> str:
        """Resolve a folder display name to its ID. Pass through if already an ID."""
        if len(name_or_id) > 50:
            return name_or_id  # likely already an ID
        # Well-known folder names work directly with the API
        well_known = {
            "inbox", "drafts", "sentitems", "deleteditems",
            "junkemail", "archive", "outbox",
        }
        if name_or_id.lower() in well_known:
            return name_or_id
        # One folder listing per client, shared safely by concurrent bulk actions.
        with self._folder_lock:
            if not self._folder_cache:
                for folder in self.get_folders():
                    self._folder_cache[folder.name.casefold()] = folder.id
            if name_or_id.casefold() in self._folder_cache:
                return self._folder_cache[name_or_id.casefold()]
        raise ResourceNotFoundError(f"Folder '{name_or_id}' not found. Run 'outlook folders' to see available folders.")

    def delete_message(self, message_id: str) -> None:
        real_id = self._resolve_id(message_id)
        self._delete(f"/messages/{real_id}")

    def mark_read(self, message_id: str, is_read: bool = True) -> None:
        real_id = self._resolve_id(message_id)
        self._patch(f"/messages/{real_id}", json={"IsRead": is_read})

    def set_flag(
        self,
        message_id: str,
        status: str = "flagged",
        due_date: str | None = None,
    ) -> dict:
        """Set the follow-up flag on a message.

        status: "flagged", "complete", or "notFlagged".
        due_date: optional ISO date string (YYYY-MM-DD) for DueDateTime/StartDateTime.
        """
        real_id = self._resolve_id(message_id)
        flag: dict = {"FlagStatus": status}
        if due_date and status == "flagged":
            flag["DueDateTime"] = {"DateTime": f"{due_date}T23:59:59", "TimeZone": "UTC"}
            flag["StartDateTime"] = {"DateTime": f"{due_date}T00:00:00", "TimeZone": "UTC"}
        return self._patch(f"/messages/{real_id}", json={"Flag": flag})

    def pin_message(self, message_id: str, pinned: bool = True) -> dict:
        """Pin or unpin a message via OWA UpdateItem with RenewTime.

        Pin sets RenewTime to a far-future date (keeps message at top).
        Unpin deletes the RenewTime field.
        """
        real_id = self._resolve_id(message_id)
        # REST v2 uses URL-safe base64 (- and _), OWA expects standard base64 (/ and +)
        real_id = real_id.replace("-", "/").replace("_", "+")

        if pinned:
            updates = [{
                "__type": "SetItemField:#Exchange",
                "Path": {
                    "__type": "PropertyUri:#Exchange",
                    "FieldURI": "RenewTime",
                },
                "Item": {
                    "__type": "Message:#Exchange",
                    "RenewTime": "4500-09-01T00:00:00.000",
                },
            }]
        else:
            updates = [{
                "__type": "DeleteItemField:#Exchange",
                "Path": {
                    "__type": "PropertyUri:#Exchange",
                    "FieldURI": "RenewTime",
                },
            }]

        return self._owa_action("UpdateItem", {
            "__type": "UpdateItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": "UTC",
                    },
                },
            },
            "Body": {
                "__type": "UpdateItemRequest:#Exchange",
                "ItemChanges": [{
                    "__type": "ItemChange:#Exchange",
                    "Updates": updates,
                    "ItemId": {
                        "__type": "ItemId:#Exchange",
                        "Id": real_id,
                    },
                }],
                "ConflictResolution": "AlwaysOverwrite",
                "MessageDisposition": "SaveOnly",
            },
        })

    # ------------------------------------------------------------------
    # Scheduled send
    # ------------------------------------------------------------------

    def schedule_send(self, to: list[str], subject: str, body: str, send_at: str,
                      cc: list[str] | None = None, html: bool = False) -> dict:
        # Creating a draft first gives cancellation an exact identity.
        draft = self.create_draft(to=to, subject=subject, body=body, cc=cc, html=html)
        return self.schedule_draft(draft.id, send_at)

    def schedule_draft(self, message_id: str, send_at: str) -> dict:
        real_id = self._resolve_id(message_id)
        msg = self._get(f"/messages/{real_id}", params={"$select": "Id,InternetMessageId,Subject,ToRecipients,CcRecipients"})
        to = [r["EmailAddress"]["Address"] for r in msg.get("ToRecipients", [])]
        cc = [r["EmailAddress"]["Address"] for r in msg.get("CcRecipients", [])]
        resp = self._patch(f"/messages/{real_id}", json={
            "SingleValueExtendedProperties": [{"PropertyId": DEFERRED_SEND_PROPERTY_ID, "Value": send_at}],
        })
        updated_id = resp.get("Id", real_id)
        # Persist before sending so a network failure never loses a potentially queued item.
        entry = self._track_scheduled(to=to, cc=cc or None, subject=msg.get("Subject", ""),
                                      send_at=send_at, message_id=updated_id,
                                      internet_message_id=resp.get("InternetMessageId") or msg.get("InternetMessageId"),
                                      status="send_pending")
        try:
            self._post(f"/messages/{updated_id}/send")
        except Exception:
            self._update_scheduled(entry["tracking_id"], status="send_unconfirmed")
            raise
        self._update_scheduled(entry["tracking_id"], status="scheduled")
        entry["status"] = "scheduled"
        return entry

    def get_scheduled_list(self) -> list[dict]:
        entries = self._load_scheduled()
        for entry in entries:
            entry["cancellable"] = bool(entry.get("message_id"))
            if not entry.get("message_id"):
                # Historical subject-only tracking cannot identify the right draft safely.
                entry["status"] = "identity_unavailable"
        return entries

    def cancel_scheduled_entry(self, index: int, expected_tracking_id: str | None = None) -> dict | None:
        # Keep the lock across delete+commit so another process cannot change the index.
        with locked_file(self._paths.scheduled_file):
            entries = self._load_scheduled()
            if index < 1 or index > len(entries):
                return None
            entry = entries[index - 1]
            if expected_tracking_id and entry.get("tracking_id") != expected_tracking_id:
                raise OutlookCliError("Schedule list changed during confirmation. Run schedule-list again; nothing was cancelled.")
            message_id = entry.get("message_id")
            if not message_id:
                raise OutlookCliError("Cannot safely cancel this legacy schedule: no message identity was recorded. Tracking preserved.")
            try:
                # Verify it is still a draft; deleting SentItems cannot recall a message.
                message = self._get(f"/messages/{message_id}", params={"$select": "Id,IsDraft,InternetMessageId"})
                if message.get("IsDraft") is not True:
                    raise OutlookCliError("Scheduled message is no longer a confirmed draft; cancellation was not performed. Tracking preserved.")
                expected = entry.get("internet_message_id")
                if expected and message.get("InternetMessageId") and message["InternetMessageId"] != expected:
                    raise OutlookCliError("Scheduled message identity changed; tracking preserved.")
                self._delete(f"/messages/{message_id}")
            except Exception as exc:
                entry["status"] = "cancellation_failed"
                entry["last_error"] = type(exc).__name__
                self._save_scheduled(entries)
                raise
            entries.pop(index - 1)
            self._save_scheduled(entries)
            return dict(entry, server_deleted=True)

    def _track_scheduled(self, to: list[str], subject: str, send_at: str,
                         cc: list[str] | None = None, message_id: str | None = None,
                         internet_message_id: str | None = None, status: str = "scheduled") -> dict:
        from datetime import datetime, timezone
        entry = {"tracking_id": str(uuid.uuid4()), "to": to, "cc": cc or [], "subject": subject,
                 "scheduled_at": send_at, "created_at": datetime.now(timezone.utc).isoformat(), "status": status}
        if message_id:
            entry["message_id"] = message_id
        if internet_message_id:
            entry["internet_message_id"] = internet_message_id
        with locked_file(self._paths.scheduled_file):
            entries = self._load_scheduled()
            entries.append(entry)
            self._save_scheduled(entries)
        return entry

    def _update_scheduled(self, tracking_id: str, **updates) -> None:
        with locked_file(self._paths.scheduled_file):
            entries = self._load_scheduled()
            for entry in entries:
                if entry.get("tracking_id") == tracking_id:
                    entry.update(updates)
                    break
            self._save_scheduled(entries)

    def _load_scheduled(self) -> list[dict]:
        if not self._paths.scheduled_file.exists():
            return []
        try:
            entries = json.loads(self._paths.scheduled_file.read_text())
            if not isinstance(entries, list) or any(not isinstance(e, dict) for e in entries):
                raise ValueError("invalid tracking file")
            return entries
        except (OSError, ValueError) as exc:
            raise AccountError("Scheduled tracking file is unreadable; preserved for repair.") from exc

    def _save_scheduled(self, entries: list[dict]) -> None:
        atomic_json_write(self._paths.scheduled_file, entries)

    def search_messages(self, query: str, top: int = 25, select: str | None = None,
                        all_pages: bool = False) -> Page:
        params = {"$search": f'"{query}"', "$top": min(top, 1000)}
        if select:
            params["$select"] = ",".join(dict.fromkeys(["Id", *select.split(",")]))
        raw = paginate(self._get, "/messages", params, limit=None if all_pages else top, search=True)
        messages = Page((Email.from_api(item) for item in raw), meta=raw.meta)
        self._assign_display_nums(messages)
        return messages

    def get_open_target(self, item_id: str) -> tuple[str, str]:
        """Resolve a display number or real ID to an Outlook on the web URL."""
        label = f"#{item_id}" if item_id.isdigit() else item_id
        try:
            real_id = self._resolve_id(item_id)
        except ResourceNotFoundError as exc:
            raise ResourceNotFoundError(
                f"Unknown item {label}. Run 'outlook inbox', 'outlook search', or 'outlook calendar' first to populate the ID map."
            ) from exc

        for kind, path in (("message", "/messages"), ("event", "/events")):
            link = self._try_get_web_link(path, real_id)
            if link:
                return kind, link

        raise ResourceNotFoundError(f"Item {label} was not found as a message or event.")

    # ------------------------------------------------------------------
    # Folders
    # ------------------------------------------------------------------

    def get_folders(self, all_pages: bool = True) -> Page:
        raw = paginate(self._get, "/MailFolders", {"$top": 100}, limit=None if all_pages else 100)
        return Page((Folder.from_api(item) for item in raw), meta=raw.meta)

    def get_folder(self, folder_id: str) -> Folder:
        resp = self._get(f"/MailFolders/{folder_id}")
        return Folder.from_api(resp)

    # ------------------------------------------------------------------
    # Attachments
    # ------------------------------------------------------------------

    def get_attachments(self, message_id: str, include_content: bool = False) -> Page:
        real_id = self._resolve_id(message_id)
        params = {} if include_content else {"$select": "Id,Name,ContentType,Size,IsInline"}
        raw = paginate(self._get, f"/messages/{real_id}/attachments", params, limit=None)
        return Page((Attachment.from_api(item) for item in raw), meta=raw.meta)

    def download_attachment(self, message_id: str, attachment_id: str) -> Attachment:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}/attachments/{attachment_id}")
        return Attachment.from_api(resp)

    def add_attachment(self, message_id: str, file_path: str) -> dict:
        """Add a file attachment to a draft message.

        Uses inline base64 for files under 3 MB, upload session for larger.
        message_id can be a display number or real Outlook ID.
        """
        import base64
        import mimetypes

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        real_id = self._resolve_id(message_id)
        file_size = path.stat().st_size

        if file_size < ATTACHMENT_SIZE_THRESHOLD:
            content = base64.b64encode(path.read_bytes()).decode()
            content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            return self._post(f"/messages/{real_id}/attachments", json={
                "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                "Name": path.name,
                "ContentType": content_type,
                "ContentBytes": content,
            })
        else:
            return self._upload_large_attachment(real_id, path, file_size)

    def _upload_large_attachment(self, real_id: str, path: Path, file_size: int) -> dict:
        """Upload a large file via an upload session (for files >= 3 MB)."""
        session = self._post(f"/messages/{real_id}/attachments/createuploadsession", json={
            "AttachmentItem": {
                "attachmentType": "file",
                "name": path.name,
                "size": file_size,
            }
        })
        upload_url = session["uploadUrl"]

        chunk_size = 4 * 1024 * 1024  # 4 MB chunks
        result: dict = {}
        with open(path, "rb") as f:
            offset = 0
            while offset < file_size:
                chunk = f.read(chunk_size)
                chunk_end = offset + len(chunk) - 1
                resp = request_response(
                    self._session("upload"), "PUT", upload_url,
                    content=chunk,
                    headers={
                        "Content-Type": "application/octet-stream",
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {offset}-{chunk_end}/{file_size}",
                    },
                    retry_safe=False,
                    account_name=self.account_name,
                )
                resp.raise_for_status()
                if resp.content:
                    result = resp.json()
                offset += len(chunk)
        return result

    def attach_files(self, message_id: str, file_paths: list[str]) -> None:
        """Attach multiple files to a draft message."""
        for fp in file_paths:
            self.add_attachment(message_id, fp)

    def create_forward_draft(self, message_id: str, to: list[str], comment: str = "") -> Email:
        """Create a forward draft without sending."""
        real_id = self._resolve_id(message_id)
        payload: dict = {
            "ToRecipients": [{"EmailAddress": {"Address": addr}} for addr in to],
        }
        data = self._post(f"/messages/{real_id}/createforward", json=payload)
        draft_id = data["Id"]

        if comment:
            # Comment field only supports plain text; patch the draft body
            # with HTML-converted content to preserve line breaks.
            original_body = data.get("Body", {}).get("Content", "")
            body_html = _plain_text_to_html(comment)
            if "<body>" in original_body:
                combined = original_body.replace("<body>", f"<body>{body_html} ", 1)
            else:
                combined = body_html + original_body
            data = self._patch(f"/messages/{draft_id}", json={
                "Body": {"ContentType": "HTML", "Content": combined},
            })

        return Email.from_api(data)

    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------

    def get_calendar_view(self, start: str, end: str, top: int = 50, calendar_name: str | None = None, all_pages: bool = False) -> list[Event]:
        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$top": top,
            "$orderby": "Start/DateTime asc",
        }
        if calendar_name:
            cal_id = self._resolve_calendar(calendar_name)
            path = f"/calendars/{cal_id}/calendarview"
        else:
            path = "/calendarview"
        raw = paginate(self._get, path, params, limit=None if all_pages else top)
        events = Page((Event.from_api(e) for e in raw), meta=raw.meta)
        self._assign_event_display_nums(events)
        return events

    def get_events(self, top: int = 25, all_pages: bool = False) -> Page:
        raw = paginate(self._get, "/events", {"$top": top, "$orderby": "Start/DateTime desc"}, limit=None if all_pages else top)
        events = Page((Event.from_api(e) for e in raw), meta=raw.meta)
        self._assign_event_display_nums(events)
        return events

    def get_event(self, event_id: str) -> Event:
        real_id = self._resolve_id(event_id)
        resp = self._get(f"/events/{real_id}")
        event = Event.from_api(resp)
        event.display_num = int(event_id) if event_id.isdigit() else 0
        return event

    def create_event(
        self,
        subject: str,
        start: str,
        end: str,
        timezone: str = "UTC",
        attendees: list[str] | None = None,
        location: str | None = None,
        body: str | None = None,
        html: bool = False,
        is_all_day: bool = False,
        reminder_minutes: int | None = 15,
        is_online_meeting: bool = False,
        recurrence: dict | None = None,
    ) -> Event:
        payload: dict = {
            "Subject": subject,
            "Start": {"DateTime": start, "TimeZone": timezone},
            "End": {"DateTime": end, "TimeZone": timezone},
            "IsAllDay": is_all_day,
        }
        if attendees:
            payload["Attendees"] = [
                {"EmailAddress": {"Address": addr}, "Type": "Required"}
                for addr in attendees
            ]
        if location:
            payload["Location"] = {"DisplayName": location}
        if body:
            content = body if html else _plain_text_to_html(body)
            payload["Body"] = {
                "ContentType": "HTML",
                "Content": content,
            }
        if reminder_minutes is not None:
            payload["IsReminderOn"] = True
            payload["ReminderMinutesBeforeStart"] = reminder_minutes
        if is_online_meeting:
            payload["IsOnlineMeeting"] = True
            payload["OnlineMeetingProvider"] = "TeamsForBusiness"
        if recurrence:
            payload["Recurrence"] = recurrence
        data = self._post("/events", json=payload)
        return Event.from_api(data)

    def get_event_instances(self, event_id: str, start: str, end: str, top: int = 50, all_pages: bool = False) -> list[Event]:
        """Get occurrences of a recurring event.

        If given an occurrence ID, resolves to its series master first.
        """
        real_id = self._resolve_id(event_id)
        # Check if this is an occurrence — need series master for /instances
        ev = self._get(f"/events/{real_id}", params={"$select": "Type,SeriesMasterId"})
        master_id = ev.get("SeriesMasterId") or real_id
        raw = paginate(self._get, f"/events/{master_id}/instances", {
            "startDateTime": start, "endDateTime": end, "$top": top,
        }, limit=None if all_pages else top)
        events = Page((Event.from_api(e) for e in raw), meta=raw.meta)
        self._assign_event_display_nums(events)
        return events

    def update_event(self, event_id: str, **kwargs) -> Event:
        """Update event fields. Accepts: subject, start, end, timezone,
        location, body, html, is_all_day, attendees (full replacement)."""
        real_id = self._resolve_id(event_id)
        payload: dict = {}
        tz = kwargs.get("timezone", "UTC")
        if "subject" in kwargs:
            payload["Subject"] = kwargs["subject"]
        if "start" in kwargs:
            payload["Start"] = {"DateTime": kwargs["start"], "TimeZone": tz}
        if "end" in kwargs:
            payload["End"] = {"DateTime": kwargs["end"], "TimeZone": tz}
        if "location" in kwargs:
            payload["Location"] = {"DisplayName": kwargs["location"]}
        if "body" in kwargs:
            body_content = kwargs["body"] if kwargs.get("html") else _plain_text_to_html(kwargs["body"])
            payload["Body"] = {
                "ContentType": "HTML",
                "Content": body_content,
            }
        if "is_all_day" in kwargs:
            payload["IsAllDay"] = kwargs["is_all_day"]
        if "attendees" in kwargs:
            payload["Attendees"] = [
                {"EmailAddress": {"Address": addr}, "Type": "Required"}
                for addr in kwargs["attendees"]
            ]
        data = self._patch(f"/events/{real_id}", json=payload)
        return Event.from_api(data)

    def add_event_attendees(self, event_id: str, new_addrs: list[str]) -> Event:
        """Add attendees to an existing event without removing current ones."""
        real_id = self._resolve_id(event_id)
        current = self._get(f"/events/{real_id}", params={"$select": "Attendees"})
        existing = current.get("Attendees", [])
        existing_addrs = {a["EmailAddress"]["Address"].lower() for a in existing}
        for addr in new_addrs:
            if addr.lower() not in existing_addrs:
                existing.append({"EmailAddress": {"Address": addr}, "Type": "Required"})
        data = self._patch(f"/events/{real_id}", json={"Attendees": existing})
        return Event.from_api(data)

    def remove_event_attendees(self, event_id: str, remove_addrs: list[str]) -> Event:
        """Remove attendees from an existing event."""
        real_id = self._resolve_id(event_id)
        current = self._get(f"/events/{real_id}", params={"$select": "Attendees"})
        existing = current.get("Attendees", [])
        remove_lower = {a.lower() for a in remove_addrs}
        filtered = [a for a in existing if a["EmailAddress"]["Address"].lower() not in remove_lower]
        data = self._patch(f"/events/{real_id}", json={"Attendees": filtered})
        return Event.from_api(data)

    def delete_event(self, event_id: str) -> None:
        real_id = self._resolve_id(event_id)
        self._delete(f"/events/{real_id}")

    def respond_to_event(self, event_id: str, response: str, comment: str = "", send_response: bool = True) -> None:
        """Respond to a meeting. response: accept, decline, tentativelyaccept."""
        real_id = self._resolve_id(event_id)
        payload = {"SendResponse": send_response}
        if comment:
            payload["Comment"] = comment
        self._post(f"/events/{real_id}/{response}", json=payload)

    def find_meeting_times(
        self,
        attendees: list[str],
        start: str,
        end: str,
        duration_minutes: int = 60,
        timezone: str = "UTC",
        max_candidates: int = 5,
    ) -> list[dict]:
        payload = {
            "Attendees": [
                {"Type": "Required", "EmailAddress": {"Address": addr}}
                for addr in attendees
            ],
            "TimeConstraint": {
                "Timeslots": [{
                    "Start": {"DateTime": start, "TimeZone": timezone},
                    "End": {"DateTime": end, "TimeZone": timezone},
                }]
            },
            "MeetingDuration": f"PT{duration_minutes}M",
            "MaxCandidates": max_candidates,
        }
        resp = self._post("/findMeetingTimes", json=payload)
        return resp.get("MeetingTimeSuggestions", [])

    def search_people(self, query: str, top: int = 10) -> list[dict]:
        resp = self._get("/people", params={"$search": query, "$top": top})
        return resp.get("value", [])

    def get_calendars(self, all_pages: bool = True) -> Page:
        return paginate(self._get, "/calendars", {"$top": 50}, limit=None if all_pages else 50)

    def _resolve_calendar(self, name: str) -> str:
        """Resolve a calendar display name to its ID."""
        cals = self.get_calendars()
        # Exact match first
        for c in cals:
            if c.get("Name", "").lower() == name.lower():
                return c["Id"]
        # Partial match
        for c in cals:
            if name.lower() in c.get("Name", "").lower():
                return c["Id"]
        available = ", ".join(c.get("Name", "") for c in cals)
        raise ResourceNotFoundError(f"Calendar '{name}' not found. Available: {available}")

    def _assign_event_display_nums(self, events: list[Event]) -> None:
        self._assign_display_nums(events)

    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------

    def get_contacts(self, top: int = 50, all_pages: bool = False) -> Page:
        raw = paginate(self._get, "/contacts", {"$top": top}, limit=None if all_pages else top)
        return Page((Contact.from_api(c) for c in raw), meta=raw.meta)

    # ------------------------------------------------------------------
    # Categories
    # ------------------------------------------------------------------

    def get_master_categories(self) -> list[dict]:
        """Fetch master category list via OWA service.svc.

        REST v2 doesn't expose /outlook/masterCategories.
        OWA uses service.svc with the action FindCategoryDetails,
        sending the JSON payload URL-encoded in the x-owa-urlpostdata header.
        """
        return self._owa_action("FindCategoryDetails", {
            "__type": "FindCategoryDetailsJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": "UTC",
                    },
                },
            },
            "Body": {
                "__type": "FindCategoryDetailsRequest:#Exchange",
            },
        })

    def get_categories(self, message_id: str) -> list[str]:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}", params={"$select": "Categories"})
        return resp.get("Categories", [])

    def set_categories(self, message_id: str, categories: list[str]) -> list[str]:
        real_id = self._resolve_id(message_id)
        resp = self._patch(f"/messages/{real_id}", json={"Categories": categories})
        return resp.get("Categories", categories)

    def add_category(self, message_id: str, category: str) -> list[str]:
        current = self.get_categories(message_id)
        if category in current:
            return current
        return self.set_categories(message_id, [*current, category])

    def remove_category(self, message_id: str, category: str) -> list[str]:
        current = self.get_categories(message_id)
        if category not in current:
            return current
        return self.set_categories(message_id, [c for c in current if c != category])

    # ------------------------------------------------------------------
    # User info
    # ------------------------------------------------------------------

    def get_me(self) -> dict:
        return self._get("")

    # ------------------------------------------------------------------
    # ID mapping
    # ------------------------------------------------------------------

    def _store(self) -> IDStore:
        # Double check under a lock: summary/bulk readers share the same client.
        with self._pool_lock:
            if self._id_store is None:
                self._id_store = IDStore(self._paths.id_map_file)
            return self._id_store

    def _resolve_id(self, display_id: str) -> str:
        display_id = str(display_id).lstrip("#")
        if display_id.isdigit():
            real_id = self._store().resolve(display_id)
            if real_id:
                return real_id
            # Compatibility for library users that supplied an in-memory mapping.
            if display_id in self._id_map:
                return self._id_map[display_id]
            raise ResourceNotFoundError(f"Unknown message #{display_id}. Run 'outlook inbox' first to populate the ID map.")
        if display_id:
            return display_id
        raise ResourceNotFoundError("Empty message ID")

    def _assign_display_nums(self, messages) -> None:
        if not messages:
            return
        numbers = self._store().allocate([item.id for item in messages])
        for item, number in zip(messages, numbers):
            item.display_num = number
        self._id_map.update({str(number): item.id for item, number in zip(messages, numbers)})

    def _evict_old_entries(self) -> None:
        """Deprecated: durable display references must never be evicted."""

    def _load_id_map(self) -> dict[str, str]:
        return self._store().snapshot()

    def _try_get_web_link(self, collection_path: str, real_id: str) -> str | None:
        """Fetch the Outlook Web URL for a message or event, if it exists."""
        try:
            resp = self._get(f"{collection_path}/{real_id}", params={"$select": "WebLink"})
        except httpx.HTTPStatusError as exc:
            if exc.response is not None and exc.response.status_code == 404:
                return None
            raise
        return resp.get("WebLink") or None

    def _save_id_map(self) -> None:
        # Allocation persists transactionally; retained for existing integrations.
        self._id_map = self._store().snapshot()

    # ------------------------------------------------------------------
    # HTTP helpers
    # ------------------------------------------------------------------

    def _get(self, path: str, params: dict | None = None) -> dict:
        return self._request("GET", path, params=params)

    def _post(self, path: str, json: dict | None = None) -> dict:
        return self._request("POST", path, json=json)

    def _patch(self, path: str, json: dict | None = None) -> dict:
        return self._request("PATCH", path, json=json)

    def _delete(self, path: str) -> dict:
        return self._request("DELETE", path)

    def _refresh_token(self):
        if self._refresh is None:
            raise TokenExpiredError("Token expired. Run: outlook login")
        token = self._refresh()
        if token:
            self._token = token
            self._client.headers["Authorization"] = f"Bearer {token}"
            if self._owa_client is not None:
                self._owa_client.headers["Authorization"] = f"Bearer {token}"
        return token

    def _request(self, method: str, path: str, params: dict | None = None,
                 json: dict | None = None, _retry: int = 0) -> dict:
        return request_json(self._client, method, path, params=params, json=json,
                            refresh=self._refresh_token if self._refresh else None,
                            account_name=self.account_name)

    def _session(self, kind: str) -> httpx.Client:
        with self._pool_lock:
            attribute = "_owa_client" if kind == "owa" else "_upload_client"
            session = getattr(self, attribute)
            if session is None:
                # Pre-authenticated upload URLs must never receive mailbox tokens.
                headers = {"User-Agent": USER_AGENT}
                if kind == "owa":
                    headers["Authorization"] = f"Bearer {self._token}"
                session = httpx.Client(headers=headers, timeout=15 if kind == "owa" else 120)
                setattr(self, attribute, session)
            return session

    def close(self) -> None:
        for session in (self._client, self._owa_client, self._upload_client):
            if session is not None:
                session.close()

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        self.close()

    def _owa_action(self, action: str, payload: dict) -> dict:
        return request_json(self._session("owa"), "POST", f"{OWA_SERVICE_URL}?action={action}",
                            headers={"Content-Type": "application/json; charset=utf-8",
                                     "Action": action, "x-req-source": "Mail",
                                     "x-owa-urlpostdata": quote(json.dumps(payload), safe="")},
                            content=b"", refresh=self._refresh_token if self._refresh else None,
                            account_name=self.account_name, retry_safe=action.startswith(("Get", "Find")))
