"""Optional, explicitly authenticated Microsoft Graph read/delta backend."""
from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from urllib.parse import quote

import httpx
import keyring

from . import account, credentials
from .exceptions import AccountError, AuthRequiredError, OutlookCliError, ResourceNotFoundError, TokenExpiredError
from .locking import file_lock
from .pagination import paginate, validated_next_link
from .transport import request_json

GRAPH_URL = "https://graph.microsoft.com/v1.0"
SERVICE = "outlook-cli-graph"
SCOPES = "https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Read offline_access"
FIELDS = "id,parentFolderId,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments,conversationId,categories,lastModifiedDateTime"


def _identity(me: dict) -> dict:
    return {"Id": me.get("id"), "EmailAddress": me.get("mail") or me.get("userPrincipalName"),
            "DisplayName": me.get("displayName", "")}


def _load_secret(profile: str, *, allow_interactive=True) -> dict:
    try:
        raw = credentials.get_password(SERVICE, profile, allow_interactive=allow_interactive)
        return json.loads(raw) if raw else {}
    except (keyring.errors.KeyringError, ValueError) as exc:
        raise AuthRequiredError("Graph credentials unavailable; run graph-login.") from exc


def _save_secret(profile: str, data: dict, *, allow_interactive=True) -> None:
    try:
        credentials.set_password(SERVICE, profile, json.dumps(data), allow_interactive=allow_interactive)
    except keyring.errors.KeyringError as exc:
        raise AuthRequiredError("Cannot save Graph credentials in the OS keychain.") from exc


def _authority(tenant: str) -> str:
    if not re.fullmatch(r"[a-zA-Z0-9.-]+", tenant):
        raise ValueError("Invalid tenant: use a tenant ID/domain or organizations/common.")
    return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"


def _verify(token: str, profile: str) -> dict:
    with httpx.Client(base_url=GRAPH_URL, headers={"Authorization": f"Bearer {token}"}, timeout=20) as client:
        me = request_json(client, "GET", "/me", params={"$select": "id,mail,userPrincipalName,displayName"}, account_name=profile)
    identity = _identity(me)
    if not identity["EmailAddress"]:
        raise AccountError("Graph did not return a mailbox identity.")
    account.assert_mailbox_matches(profile, identity)
    return identity


def _verify_and_bind(token: str, profile: str) -> dict:
    identity = _verify(token, profile)
    # Recheck identity under the registry lock even when another login bound it
    # between verification and this write.
    account.bind_account(profile, identity)
    return identity


def graph_login(profile: str, client_id: str, tenant: str = "organizations", timeout: int = 300) -> dict:
    """Explicit device flow only; never called implicitly by a read operation."""
    try:
        import uuid
        uuid.UUID(client_id)
    except ValueError as exc:
        raise ValueError("client-id must be the Entra application's UUID.") from exc
    authority = _authority(tenant)
    lock = account.get_account_paths(profile).cache_dir / "graph-auth.lock"
    with file_lock(lock, timeout=timeout + 30), httpx.Client(timeout=20) as client:
        response = client.post(authority + "/devicecode", data={"client_id": client_id, "scope": SCOPES})
        response.raise_for_status()
        flow = response.json()
        print(flow.get("message") or f"Open {flow['verification_uri']} and enter {flow['user_code']}", file=sys.stderr)
        deadline = time.monotonic() + min(timeout, int(flow.get("expires_in", timeout)))
        interval = max(1, int(flow.get("interval", 5)))
        while time.monotonic() < deadline:
            time.sleep(min(interval, max(0, deadline - time.monotonic())))
            if time.monotonic() >= deadline:
                break
            response = client.post(authority + "/token", data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id, "device_code": flow["device_code"],
            })
            data = response.json()
            error = data.get("error")
            if error == "authorization_pending":
                continue
            if error == "slow_down":
                interval += 5
                continue
            if error:
                raise AuthRequiredError(f"Graph login failed: {error}. Check app permissions/consent.")
            response.raise_for_status()
            token = data.get("access_token")
            if not token:
                raise AuthRequiredError("Graph login did not return an access token.")
            _verify_and_bind(token, profile)
            data.update(client_id=client_id, tenant=tenant, expires_at=time.time() + int(data.get("expires_in", 3600)))
            _save_secret(profile, data)
            return {"account": profile, "backend": "graph", "authenticated": True,
                    "expires_at": datetime.fromtimestamp(data["expires_at"], timezone.utc).isoformat()}
    raise AuthRequiredError("Graph device login timed out. Run graph-login again when ready.")


def get_graph_token(profile: str, *, force: bool = False, rejected_token: str | None = None, allow_interactive=True) -> str:
    env = os.environ.get("OUTLOOK_GRAPH_TOKEN")
    if env:
        if force:
            raise AuthRequiredError("OUTLOOK_GRAPH_TOKEN was rejected; provide a fresh Graph token.")
        _verify_and_bind(env, profile)
        return env
    lock = account.get_account_paths(profile).cache_dir / "graph-auth.lock"
    with file_lock(lock, timeout=30):
        data = _load_secret(profile, allow_interactive=allow_interactive)
        token = data.get("access_token")
        valid = data.get("expires_at", 0) > time.time() + 120
        if token and valid and (not force or token != rejected_token):
            try:
                _verify_and_bind(token, profile)
                return token
            except TokenExpiredError:
                pass  # Revoked tokens can expire before the cached TTL; refresh once.
        if not data.get("refresh_token") or not data.get("client_id"):
            raise AuthRequiredError("Graph is not configured. Run graph-login --client-id APP_ID, or use --backend rest.")
        with httpx.Client(timeout=20) as client:
            response = client.post(_authority(data.get("tenant", "organizations")) + "/token", data={
                "grant_type": "refresh_token", "client_id": data["client_id"],
                "refresh_token": data["refresh_token"], "scope": SCOPES,
            })
        refreshed = response.json()
        if response.status_code != 200 or not refreshed.get("access_token"):
            raise AuthRequiredError("Graph session needs interactive authentication; run graph-login.")
        _verify_and_bind(refreshed["access_token"], profile)
        data.update(refreshed)
        data["expires_at"] = time.time() + int(refreshed.get("expires_in", 3600))
        _save_secret(profile, data, allow_interactive=allow_interactive)
        return data["access_token"]


class GraphReader:
    """Read-only adapter. Stable IDs and delta state are scoped to an account."""
    def __init__(self, token: str, profile: str = "default", refresh=None):
        self.profile = profile
        self.token = token
        self.refresh = refresh
        self.client = httpx.Client(base_url=GRAPH_URL, headers={
            "Authorization": f"Bearer {token}", "Prefer": 'IdType="ImmutableId", outlook.body-content-type="text"',
        }, timeout=30)

    @classmethod
    def for_account(cls, profile: str):
        from .commands._common import is_no_input_mode
        allow_interactive = not is_no_input_mode()
        token = get_graph_token(profile, allow_interactive=allow_interactive)
        reader = cls(token, profile)
        def refresh():
            reader.token = get_graph_token(profile, force=True, rejected_token=reader.token, allow_interactive=allow_interactive)
            return reader.token
        reader.refresh = refresh
        return reader

    def close(self):
        self.client.close()

    def get(self, path: str, params=None):
        # Validate the final httpx URL, including uppercase schemes and //host paths.
        target = str(self.client.build_request("GET", path).url)
        validated_next_link(target, GRAPH_URL + "/me", GRAPH_URL)
        return request_json(self.client, "GET", path, params=params, refresh=self.refresh, account_name=self.profile)

    def folders(self):
        queue, result, seen = ["/me/mailFolders"], [], set()
        while queue:
            if len(seen) > 10000:
                raise OutlookCliError("Folder hierarchy exceeded safety limit.")
            path = queue.pop(0)
            page = paginate(self.get, path, {"$top": 100, "includeHiddenFolders": "true"}, limit=None, base_url=GRAPH_URL)
            if not page.meta.get("complete"):
                raise OutlookCliError("Folder hierarchy is incomplete; retry with narrower folder selection.")
            for folder in page:
                if folder["id"] in seen:
                    continue
                seen.add(folder["id"])
                result.append(folder)
                if folder.get("childFolderCount", 0):
                    queue.append(f"/me/mailFolders/{quote(folder['id'], safe='')}/childFolders")
        return result

    def messages(self, folder_id: str, *, include_body=False):
        return paginate(self.get, f"/me/mailFolders/{quote(folder_id, safe='')}/messages",
            {"$top": 100, "$select": FIELDS + (",body" if include_body else "")},
            limit=None, base_url=GRAPH_URL)

    def delta(self, folder_id: str, cursor: str | None = None, *, include_body=False, max_pages=1000):
        path = cursor or f"/me/mailFolders/{quote(folder_id, safe='')}/messages/delta"
        params = None if cursor else {"$select": FIELDS + (",body" if include_body else ""), "$top": 100}
        rows, links = [], set()
        for _ in range(max_pages):
            response = self.get(path, params=params)
            values = response.get("value")
            if not isinstance(values, list):
                raise OutlookCliError("Graph delta response is missing its value collection.")
            rows.extend(values)
            next_link = response.get("@odata.nextLink")
            if not next_link:
                delta_link = response.get("@odata.deltaLink")
                if not delta_link:
                    raise OutlookCliError("Graph delta did not return a checkpoint; index was not changed.")
                validated_next_link(delta_link, GRAPH_URL + "/me", GRAPH_URL)
                return rows, delta_link
            validated_next_link(next_link, GRAPH_URL + "/me", GRAPH_URL)
            if next_link in links:
                raise OutlookCliError("Graph returned a repeated delta cursor; index was not changed.")
            links.add(next_link)
            path, params = next_link, None
        raise OutlookCliError("Graph delta exceeded its page budget; previous checkpoint retained.")


def materialize_changes(reader, changes: list[dict], folder_id: str, *, include_body=False) -> list[dict]:
    """Resolve repeated/partial delta entries without assuming response order."""
    entries, repeated = {}, set()
    for message in changes:
        identity = message["id"]
        if identity in entries:
            repeated.add(identity)
        entries[identity] = message
    fields = FIELDS + (",body" if include_body else "")
    required = set(fields.split(","))
    result = []
    for identity, message in entries.items():
        if identity in repeated or ("@removed" not in message and not required.issubset(message)):
            try:
                message = reader.get(f"/me/messages/{quote(identity, safe='')}", params={"$select": fields})
            except (httpx.HTTPStatusError, ResourceNotFoundError) as exc:
                if not isinstance(exc, ResourceNotFoundError) and exc.response.status_code != 404:
                    raise
                message = {"id": identity, "@removed": {"reason": "deleted"}}
            if "@removed" not in message and (message.get("id") != identity or not message.get("parentFolderId")):
                raise OutlookCliError("Graph message lookup returned an incomplete identity; checkpoint retained.")
        if "@removed" not in message and message.get("parentFolderId") != folder_id:
            message = {"id": identity, "@removed": {"reason": "moved"}}
        result.append(message)
    return result


def graph_to_record(message: dict) -> dict:
    def addresses(items):
        return [{"name": (x.get("emailAddress") or {}).get("name", ""),
                 "address": (x.get("emailAddress") or {}).get("address", "")} for x in items or []]
    sender = addresses([message.get("from") or {}])[0]
    body = message.get("body") or {}
    return {"id": message["id"], "subject": message.get("subject", ""), "sender": sender,
            "to": addresses(message.get("toRecipients", [])), "cc": addresses(message.get("ccRecipients", [])),
            "received": message.get("receivedDateTime", ""), "preview": message.get("bodyPreview", ""),
            "body": body.get("content", ""), "body_type": body.get("contentType", "text"),
            "conversation_id": message.get("conversationId", ""), "categories": message.get("categories") or [],
            "is_read": message.get("isRead", False), "has_attachments": message.get("hasAttachments", False)}
