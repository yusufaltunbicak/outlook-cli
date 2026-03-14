from __future__ import annotations

import json
import time
from pathlib import Path

import httpx

from urllib.parse import quote

from .constants import ATTACHMENT_SIZE_THRESHOLD, BASE_URL, CACHE_DIR, DEFERRED_SEND_PROPERTY_ID, ID_MAP_FILE, OWA_SERVICE_URL, SCHEDULED_FILE, USER_AGENT
from .exceptions import RateLimitError, ResourceNotFoundError, TokenExpiredError
from .models import Attachment, Contact, Email, Event, Folder


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
        filter_parts.append(f"Categories/any(c:c eq '{filter_category}')")
    return " and ".join(filter_parts), "", False


class OutlookClient:
    """HTTP client for Outlook REST API v2."""

    MAX_ID_MAP_SIZE = 500

    def __init__(self, token: str):
        self._token = token
        self._client = httpx.Client(
            base_url=BASE_URL,
            headers={
                "Authorization": f"Bearer {token}",
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json",
            },
            timeout=30,
        )
        self._id_map: dict[str, str] = self._load_id_map()
        self._next_num: int = max((int(k) for k in self._id_map if k.isdigit()), default=0) + 1

    # ------------------------------------------------------------------
    # Mail
    # ------------------------------------------------------------------

    def get_messages(
        self,
        folder: str = "Inbox",
        top: int = 25,
        skip: int = 0,
        unread_only: bool = False,
        filter_from: str | None = None,
        filter_subject: str | None = None,
        filter_after: str | None = None,
        filter_before: str | None = None,
        filter_has_attachments: bool = False,
        filter_category: str | None = None,
        filter_no_category: bool = False,
        select: str | None = None,
    ) -> list[Email]:
        filter_str, search_str, needs_search = _build_query_params(
            unread_only=unread_only,
            filter_from=filter_from,
            filter_subject=filter_subject,
            filter_after=filter_after,
            filter_before=filter_before,
            filter_has_attachments=filter_has_attachments,
            filter_category=filter_category,
        )

        if not filter_no_category:
            # Standard fetch — server-side filtering is sufficient
            if needs_search:
                params: dict = {"$top": top, "$search": search_str}
                if select:
                    params["$select"] = select
                resp = self._get("/messages", params=params)
            else:
                params = {
                    "$top": top,
                    "$skip": skip,
                    "$orderby": "ReceivedDateTime desc",
                }
                if filter_str:
                    params["$filter"] = filter_str
                if select:
                    params["$select"] = select
                resp = self._get(f"/MailFolders/{folder}/messages", params=params)
            messages = [Email.from_api(m) for m in resp.get("value", [])]
        else:
            # Client-side filtering: over-fetch in pages until we have enough
            messages: list[Email] = []
            batch_size = top * 3  # fetch 3x to compensate for filtering
            current_skip = skip
            max_pages = 5  # safety limit
            for _ in range(max_pages):
                if needs_search:
                    params = {"$top": batch_size, "$search": search_str}
                    if select:
                        params["$select"] = select
                    resp = self._get("/messages", params=params)
                else:
                    params = {
                        "$top": batch_size,
                        "$skip": current_skip,
                        "$orderby": "ReceivedDateTime desc",
                    }
                    if filter_str:
                        params["$filter"] = filter_str
                    if select:
                        params["$select"] = select
                    resp = self._get(f"/MailFolders/{folder}/messages", params=params)
                batch = resp.get("value", [])
                if not batch:
                    break
                for m in batch:
                    email = Email.from_api(m)
                    if not email.categories:
                        messages.append(email)
                        if len(messages) >= top:
                            break
                if len(messages) >= top or len(batch) < batch_size:
                    break
                current_skip += batch_size
            messages = messages[:top]

        self._assign_display_nums(messages)
        return messages

    def get_message(self, message_id: str) -> Email:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}")
        email = Email.from_api(resp)
        email.display_num = int(message_id) if message_id.isdigit() else 0
        return email

    def get_thread(self, message_id: str, max_messages: int = 50) -> list[Email]:
        """Fetch all messages in the same conversation as the given message.

        REST v2 doesn't support $filter on ConversationId, so we search by
        the base subject (strip Re:/Fwd: prefixes) and then filter client-side
        by matching ConversationId. Results are sorted oldest-first.
        """
        import re

        email = self.get_message(message_id)
        conv_id = email.conversation_id
        if not conv_id:
            return [email]

        # Strip reply/forward prefixes to get the base subject for search
        base_subject = re.sub(
            r'^(Re|Fwd|İlt|Ynt|Fw|AW|SV|VS)\s*:\s*',
            '', email.subject, flags=re.IGNORECASE,
        ).strip()

        if not base_subject:
            return [email]

        # Search by subject, then filter by ConversationId client-side
        resp = self._get(
            "/messages",
            params={
                "$search": f'"subject:{base_subject}"',
                "$top": max_messages,
            },
        )
        all_msgs = [Email.from_api(m) for m in resp.get("value", [])]
        thread = [m for m in all_msgs if m.conversation_id == conv_id]

        # Sort chronologically (oldest first)
        thread.sort(key=lambda m: m.received)

        self._assign_display_nums(thread)
        return thread if thread else [email]

    def send_mail(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
        send_at: str | None = None,
    ) -> None:
        message: dict = {
            "Subject": subject,
            "Body": {
                "ContentType": "HTML" if html else "Text",
                "Content": body,
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
        payload: dict = {
            "Subject": subject,
            "Body": {
                "ContentType": "HTML" if html else "Text",
                "Content": body,
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
        real_id = self._resolve_id(message_id)
        action = "replyall" if reply_all else "reply"
        self._post(f"/messages/{real_id}/{action}", json={"Comment": comment})

    def create_reply_draft(
        self,
        message_id: str,
        comment: str = "",
        reply_all: bool = False,
        html: bool = False,
    ) -> Email:
        real_id = self._resolve_id(message_id)
        action = "createreplyall" if reply_all else "createreply"

        if html and comment:
            # createReply only supports plain-text Comment.
            # Create empty draft first, then prepend HTML body before quoted reply.
            data = self._post(f"/messages/{real_id}/{action}", json={})
            draft_id = data["Id"]
            original_body = data.get("Body", {}).get("Content", "")
            # Insert user's HTML before the quoted original message
            if "<body>" in original_body:
                combined = original_body.replace("<body>", f"<body>{comment} ", 1)
            else:
                combined = comment + original_body
            data = self._patch(f"/messages/{draft_id}", json={
                "Body": {"ContentType": "HTML", "Content": combined},
            })
        else:
            payload = {}
            if comment:
                payload["Comment"] = comment
            data = self._post(f"/messages/{real_id}/{action}", json=payload)

        return Email.from_api(data)

    def forward(self, message_id: str, to: list[str], comment: str = "") -> None:
        real_id = self._resolve_id(message_id)
        payload = {
            "Comment": comment,
            "ToRecipients": [{"EmailAddress": {"Address": addr}} for addr in to],
        }
        self._post(f"/messages/{real_id}/forward", json=payload)

    def move_message(self, message_id: str, destination_folder: str) -> Email:
        real_id = self._resolve_id(message_id)
        folder_id = self._resolve_folder(destination_folder)
        resp = self._post(
            f"/messages/{real_id}/move",
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
        # Search by display name
        folders = self.get_folders()
        for f in folders:
            if f.name.lower() == name_or_id.lower():
                return f.id
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

    def schedule_send(
        self,
        to: list[str],
        subject: str,
        body: str,
        send_at: str,
        cc: list[str] | None = None,
        html: bool = False,
    ) -> dict:
        """Schedule an email via /sendmail with deferred send time.

        send_at must be ISO 8601 format (e.g. 2024-03-15T10:00:00Z).
        Returns the tracked schedule entry.
        """
        self.send_mail(to=to, subject=subject, body=body, cc=cc, html=html, send_at=send_at)
        entry = self._track_scheduled(
            to=to, cc=cc, subject=subject, send_at=send_at,
        )
        return entry

    def schedule_draft(self, message_id: str, send_at: str) -> dict:
        """Set deferred send time on an existing draft and send it."""
        real_id = self._resolve_id(message_id)
        # Read draft details for tracking
        msg = self._get(f"/messages/{real_id}", params={"$select": "Subject,ToRecipients,CcRecipients"})
        to = [r["EmailAddress"]["Address"] for r in msg.get("ToRecipients", [])]
        cc = [r["EmailAddress"]["Address"] for r in msg.get("CcRecipients", [])]
        subject = msg.get("Subject", "")

        resp = self._patch(f"/messages/{real_id}", json={
            "SingleValueExtendedProperties": [{
                "PropertyId": DEFERRED_SEND_PROPERTY_ID,
                "Value": send_at,
            }],
        })
        updated_id = resp.get("Id", real_id)
        self._post(f"/messages/{updated_id}/send")

        entry = self._track_scheduled(
            to=to, cc=cc or None, subject=subject, send_at=send_at,
            message_id=updated_id,
        )
        return entry

    def get_scheduled_list(self) -> list[dict]:
        """Get scheduled messages from local tracking + Drafts cross-check.

        REST v2 doesn't support $filter/$expand on extended properties,
        so we cross-reference local tracking with Drafts folder by subject
        to enrich entries with message_id for server-side cancellation.
        """
        entries = self._load_scheduled()
        if not entries:
            return entries

        # Cross-check with Drafts to find matching message IDs
        try:
            resp = self._get("/MailFolders/Drafts/messages", params={
                "$top": 50,
                "$select": "Id,Subject",
            })
            drafts_by_subject: dict[str, str] = {}
            for m in resp.get("value", []):
                drafts_by_subject[m.get("Subject", "")] = m["Id"]

            for entry in entries:
                if not entry.get("message_id"):
                    draft_id = drafts_by_subject.get(entry.get("subject", ""))
                    if draft_id:
                        entry["message_id"] = draft_id
        except Exception:
            pass  # server unavailable, show local-only

        return entries

    def cancel_scheduled_entry(self, index: int) -> dict | None:
        """Cancel a scheduled entry by its 1-based index.

        If a matching draft is found on server, deletes it.
        Always removes from local tracking.
        """
        # Use enriched list (with message_id from Drafts cross-check)
        enriched = self.get_scheduled_list()
        if index < 1 or index > len(enriched):
            return None

        removed = enriched[index - 1]

        # Remove from local tracking
        local = self._load_scheduled()
        if index - 1 < len(local):
            local.pop(index - 1)
            self._save_scheduled(local)

        # Try to delete the draft from server
        msg_id = removed.get("message_id")
        if msg_id:
            try:
                self._delete(f"/messages/{msg_id}")
                removed["server_deleted"] = True
            except Exception:
                removed["server_deleted"] = False

        return removed

    # ------------------------------------------------------------------
    # Scheduled tracking helpers
    # ------------------------------------------------------------------

    def _track_scheduled(
        self,
        to: list[str],
        subject: str,
        send_at: str,
        cc: list[str] | None = None,
        message_id: str | None = None,
    ) -> dict:
        from datetime import datetime as dt, timezone as tz
        entry = {
            "to": to,
            "cc": cc or [],
            "subject": subject,
            "scheduled_at": send_at,
            "created_at": dt.now(tz.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        if message_id:
            entry["message_id"] = message_id
        entries = self._load_scheduled()
        entries.append(entry)
        self._save_scheduled(entries)
        return entry

    def _load_scheduled(self) -> list[dict]:
        if SCHEDULED_FILE.exists():
            try:
                return json.loads(SCHEDULED_FILE.read_text())
            except (json.JSONDecodeError, OSError):
                pass
        return []

    def _save_scheduled(self, entries: list[dict]) -> None:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        SCHEDULED_FILE.write_text(json.dumps(entries, indent=2))

    def search_messages(self, query: str, top: int = 25) -> list[Email]:
        params = {
            "$search": f'"{query}"',
            "$top": top,
        }
        resp = self._get("/messages", params=params)
        messages = [Email.from_api(m) for m in resp.get("value", [])]
        self._assign_display_nums(messages)
        return messages

    # ------------------------------------------------------------------
    # Folders
    # ------------------------------------------------------------------

    def get_folders(self) -> list[Folder]:
        resp = self._get("/MailFolders", params={"$top": 100})
        return [Folder.from_api(f) for f in resp.get("value", [])]

    def get_folder(self, folder_id: str) -> Folder:
        resp = self._get(f"/MailFolders/{folder_id}")
        return Folder.from_api(resp)

    # ------------------------------------------------------------------
    # Attachments
    # ------------------------------------------------------------------

    def get_attachments(self, message_id: str) -> list[Attachment]:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}/attachments")
        return [Attachment.from_api(a) for a in resp.get("value", [])]

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
                resp = httpx.put(
                    upload_url,
                    content=chunk,
                    headers={
                        "Content-Type": "application/octet-stream",
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {offset}-{chunk_end}/{file_size}",
                    },
                    timeout=120,
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
        if comment:
            payload["Comment"] = comment
        data = self._post(f"/messages/{real_id}/createforward", json=payload)
        return Email.from_api(data)

    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------

    def get_calendar_view(self, start: str, end: str, top: int = 50, calendar_name: str | None = None) -> list[Event]:
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
        resp = self._get(path, params=params)
        events = [Event.from_api(e) for e in resp.get("value", [])]
        self._assign_event_display_nums(events)
        return events

    def get_events(self, top: int = 25) -> list[Event]:
        resp = self._get("/events", params={"$top": top, "$orderby": "Start/DateTime desc"})
        events = [Event.from_api(e) for e in resp.get("value", [])]
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
            payload["Body"] = {
                "ContentType": "HTML" if html else "Text",
                "Content": body,
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

    def get_event_instances(self, event_id: str, start: str, end: str, top: int = 50) -> list[Event]:
        """Get occurrences of a recurring event.

        If given an occurrence ID, resolves to its series master first.
        """
        real_id = self._resolve_id(event_id)
        # Check if this is an occurrence — need series master for /instances
        ev = self._get(f"/events/{real_id}", params={"$select": "Type,SeriesMasterId"})
        master_id = ev.get("SeriesMasterId") or real_id
        resp = self._get(f"/events/{master_id}/instances", params={
            "startDateTime": start,
            "endDateTime": end,
            "$top": top,
        })
        events = [Event.from_api(e) for e in resp.get("value", [])]
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
            payload["Body"] = {
                "ContentType": "HTML" if kwargs.get("html") else "Text",
                "Content": kwargs["body"],
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

    def get_calendars(self) -> list[dict]:
        resp = self._get("/calendars", params={"$top": 50})
        return resp.get("value", [])

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
        """Assign display numbers to events using the shared ID map."""
        for ev in events:
            existing = next(
                (k for k, v in self._id_map.items() if v == ev.id and k.isdigit()),
                None,
            )
            if existing:
                ev.display_num = int(existing)
            else:
                ev.display_num = self._next_num
                self._id_map[str(self._next_num)] = ev.id
                self._next_num += 1
        self._evict_old_entries()
        self._save_id_map()

    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------

    def get_contacts(self, top: int = 50) -> list[Contact]:
        resp = self._get("/contacts", params={"$top": top})
        return [Contact.from_api(c) for c in resp.get("value", [])]

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
        if category not in current:
            current.append(category)
        return self.set_categories(message_id, current)

    def remove_category(self, message_id: str, category: str) -> list[str]:
        current = self.get_categories(message_id)
        current = [c for c in current if c != category]
        return self.set_categories(message_id, current)

    # ------------------------------------------------------------------
    # User info
    # ------------------------------------------------------------------

    def get_me(self) -> dict:
        return self._get("")

    # ------------------------------------------------------------------
    # ID mapping
    # ------------------------------------------------------------------

    def _resolve_id(self, display_id: str) -> str:
        """Convert display number to real Outlook ID."""
        if display_id in self._id_map:
            return self._id_map[display_id]
        # Maybe it's already a real ID (long base64 string)
        if len(display_id) > 50:
            return display_id
        raise ResourceNotFoundError(
            f"Unknown message #{display_id}. Run 'outlook inbox' first to populate the ID map."
        )

    def _assign_display_nums(self, messages: list[Email]) -> None:
        for msg in messages:
            # Check if this real ID already has a display number
            existing = next(
                (k for k, v in self._id_map.items() if v == msg.id and k.isdigit()),
                None,
            )
            if existing:
                msg.display_num = int(existing)
            else:
                msg.display_num = self._next_num
                self._id_map[str(self._next_num)] = msg.id
                self._next_num += 1
        self._evict_old_entries()
        self._save_id_map()

    def _evict_old_entries(self) -> None:
        """Keep only the most recent MAX_ID_MAP_SIZE entries."""
        numeric = sorted(
            ((int(k), k) for k in self._id_map if k.isdigit()),
            key=lambda x: x[0],
        )
        if len(numeric) <= self.MAX_ID_MAP_SIZE:
            return
        to_remove = numeric[: len(numeric) - self.MAX_ID_MAP_SIZE]
        for _, k in to_remove:
            del self._id_map[k]

    def _load_id_map(self) -> dict[str, str]:
        if ID_MAP_FILE.exists():
            try:
                return json.loads(ID_MAP_FILE.read_text())
            except (json.JSONDecodeError, OSError):
                pass
        return {}

    def _save_id_map(self) -> None:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        ID_MAP_FILE.write_text(json.dumps(self._id_map))

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

    def _request(
        self,
        method: str,
        path: str,
        params: dict | None = None,
        json: dict | None = None,
        _retry: int = 0,
    ) -> dict:
        resp = self._client.request(method, path, params=params, json=json)

        if resp.status_code == 401:
            raise TokenExpiredError("Token expired. Run: outlook login")

        if resp.status_code == 429:
            if _retry >= 3:
                raise RateLimitError("Rate limited after 3 retries")
            retry_after = int(resp.headers.get("Retry-After", 2 ** (_retry + 1)))
            time.sleep(retry_after)
            return self._request(method, path, params=params, json=json, _retry=_retry + 1)

        if resp.status_code == 204:
            return {}

        resp.raise_for_status()

        if not resp.content:
            return {}
        return resp.json()


    def _owa_action(self, action: str, payload: dict) -> dict:
        """Call OWA service.svc endpoint.

        OWA uses a non-standard pattern: the JSON payload is URL-encoded
        in the x-owa-urlpostdata header, and the body is empty.
        """
        resp = httpx.post(
            f"{OWA_SERVICE_URL}?action={action}",
            headers={
                "Authorization": f"Bearer {self._token}",
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json; charset=utf-8",
                "Action": action,
                "x-req-source": "Mail",
                "x-owa-urlpostdata": quote(json.dumps(payload), safe=""),
            },
            content=b"",
            timeout=15,
        )
        if resp.status_code == 401:
            raise TokenExpiredError("Token expired. Run: outlook login")
        resp.raise_for_status()
        return resp.json()


