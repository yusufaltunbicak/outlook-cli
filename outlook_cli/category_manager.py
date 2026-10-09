"""Manage Outlook master categories via OWA service.svc API.

Uses the UpdateMasterCategoryList action with a request-wrapped payload.
Key discovery: the payload must be wrapped in {"request": {...}} and sent
via x-owa-urlpostdata header (not in the body).

Operations:
  - Create: AddCategoryList
  - Delete: RemoveCategoryList
  - Rename: RemoveCategoryList + AddCategoryList (with same Id, new Name)
  - Recolor: ChangeCategoryColorList
"""

from __future__ import annotations

import json
import time
import hashlib
from contextlib import contextmanager
from pathlib import Path

import click
import uuid
from datetime import datetime, timezone
from typing import Callable
from urllib.parse import quote

import httpx

from .constants import BASE_URL, USER_AGENT
from .exceptions import ResourceNotFoundError, TokenExpiredError, PartialFailureError
from . import account as account_service
from .locking import file_lock, atomic_write_json
from .transport import request_json

OWA_SERVICE_BASE = "https://outlook.cloud.microsoft/owa/service.svc"


def bind_client(client) -> None:
    """Reuse an invocation's HTTP pools and token refresh for standalone helpers."""
    ctx = click.get_current_context(silent=True)
    if ctx is not None:
        ctx.meta["outlook_manager_client"] = client


def _request(client, method, path, **kwargs):
    bound = _bound_client()
    return request_json(client, method, path, account_name=getattr(bound, "account_name", None), **kwargs)


def _bound_client():
    ctx = click.get_current_context(silent=True)
    return ctx.meta.get("outlook_manager_client") if ctx is not None else None


@contextmanager
def _session(token: str, *, owa=False):
    bound = _bound_client()
    if bound is not None:
        client = bound._session("owa") if owa else bound._client
        yield client, getattr(bound, "_refresh_token", None)
    else:
        client = httpx.Client(headers={"Authorization": f"Bearer {token}",
            "Content-Type": "application/json", "User-Agent": USER_AGENT}, timeout=30)
        try:
            yield client, None
        finally:
            client.close()


def _owa_request(token: str, action: str, payload: dict) -> dict:
    """Use the shared retry/refresh policy for OWA calls."""
    with _session(token, owa=True) as (client, refresh):
        return _request(client, "POST", f"{OWA_SERVICE_BASE}?action={action}",
            headers={"Content-Type": "application/json; charset=utf-8", "Action": action,
                "x-req-source": "Mail", "x-owa-urlpostdata": quote(json.dumps(payload), safe="")},
            content=b"", refresh=refresh, retry_safe=action == "GetOwaUserConfiguration")


def _update_master_categories(
    token: str,
    add: list[dict] | None = None,
    remove: list[str] | None = None,
    change_color: list[dict] | None = None,
) -> dict:
    """Call UpdateMasterCategoryList with the request-wrapped payload."""
    payload = {
        "request": {
            "__type": "UpdateMasterCategoryListRequest:#Exchange",
            "AddCategoryList": add or [],
            "RemoveCategoryList": remove or [],
            "ChangeCategoryColorList": change_color or [],
            "UpdateCategoryLastTimeUsedList": [],
            "ChangeCategoryKeyboardShortcutList": [],
        }
    }
    return _owa_request(token, "UpdateMasterCategoryList", payload)


def get_master_categories(token: str) -> list[dict]:
    """Fetch master category list via GetOwaUserConfiguration."""
    payload = {
        "__type": "GetOwaUserConfigurationJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2018_01_08",
        },
        "Body": {
            "__type": "GetOwaUserConfigurationRequest:#Exchange",
            "Owaconfigs": ["MasterCategoryList"],
        },
    }
    resp = _owa_request(token, "GetOwaUserConfiguration", payload)
    return resp.get("MasterCategoryList", {}).get("MasterList", [])


def create_category(token: str, name: str, color: int = 15) -> dict:
    """Create a new master category."""
    cat = {
        "Name": name,
        "Color": color,
        "Id": str(uuid.uuid4()),
        "LastTimeUsed": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "KeyboardShortcut": 0,
    }
    return _update_master_categories(token, add=[cat])


def delete_category(token: str, name: str) -> dict:
    """Delete a master category by name."""
    return _update_master_categories(token, remove=[name])


def rename_category(
    token: str,
    old_name: str,
    new_name: str,
    propagate: bool = True,
    on_progress: Callable[[int, int], None] | None = None,
) -> int:
    """Rename a master category and optionally propagate to all messages.

    Returns the number of messages updated.
    """
    master = get_master_categories(token)
    existing = next((c for c in master if c["Name"] == old_name), None)
    if not existing:
        checkpoint = _checkpoint_path(old_name, new_name, None)
        if propagate and checkpoint.exists() and any(c["Name"] == new_name for c in master):
            return _bulk_rename_on_messages(token, old_name, new_name, on_progress)
        raise ResourceNotFoundError(f"Category '{old_name}' not found.")

    new_cat = {
        **existing,
        "Name": new_name,
        "LastTimeUsed": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
    }
    _update_master_categories(token, add=[new_cat], remove=[old_name])

    if not propagate:
        return 0

    return _bulk_rename_on_messages(token, old_name, new_name, on_progress)


def _checkpoint_path(name: str, replacement: str | None, folder: str | None) -> Path:
    bound = _bound_client()
    selected = account_service.resolve_account_name(getattr(bound, "account_name", None))
    operation = json.dumps([name, replacement, folder], ensure_ascii=False)
    digest = hashlib.sha256(operation.encode()).hexdigest()[:20]
    return account_service.get_account_paths(selected).cache_dir / "operations" / f"category-{digest}.json"


def _bulk_rename_on_messages(token, old_name, new_name, on_progress=None) -> int:
    return _bulk_categories(token, old_name, new_name, on_progress=on_progress)


def clear_category(token, name, folder=None, max_messages=None, on_progress=None) -> int:
    """Clear labels with finite retries and a durable, automatic resume checkpoint."""
    return _bulk_categories(token, name, None, folder=folder,
        max_messages=max_messages, on_progress=on_progress)


def _bulk_categories(token, name, replacement, *, folder=None, max_messages=None, on_progress=None):
    checkpoint = _checkpoint_path(name, replacement, folder)
    with file_lock(checkpoint.with_suffix(".lock")):
        state = json.loads(checkpoint.read_text()) if checkpoint.exists() else {
            "name": name, "replacement": replacement, "folder": folder, "completed": [], "failures": []}
        completed = set(state.get("completed", []))
        initial_count = len(completed)
        failures = []

        def persist():
            state.update(completed=sorted(completed), failures=failures)
            atomic_write_json(checkpoint, state)

        def stopped(reason):
            persist()
            raise PartialFailureError(f"{reason} Updated {len(completed)} messages. "
                f"Repeat the same command to resume. Checkpoint: {checkpoint}",
                completed=len(completed), failures=failures, checkpoint=checkpoint)

        persist()
        with _session(token) as (client, refresh):
            path = f"{BASE_URL}/MailFolders/{quote(folder, safe='')}/messages" if folder else f"{BASE_URL}/messages"
            escaped = name.replace("'", "''")
            for page in range(1000):
                try:
                    data = _request(client, "GET", path, params={"$top": 50,
                        "$filter": f"Categories/any(c:c eq '{escaped}')", "$select": "Id,Categories"}, refresh=refresh)
                except Exception as exc:
                    failures.append({"id": None, "message": str(exc), "phase": "list"})
                    stopped("Category enumeration failed.")
                messages = data.get("value", [])
                if not messages:
                    checkpoint.unlink(missing_ok=True)
                    return len(completed)
                progress = 0
                for message in messages:
                    item_id = message["Id"]
                    if item_id in completed:
                        continue
                    categories = list(dict.fromkeys(replacement if category == name else category
                        for category in message.get("Categories", []) if replacement is not None or category != name))
                    try:
                        _request(client, "PATCH", f"{BASE_URL}/messages/{quote(item_id, safe='')}",
                            json={"Categories": categories}, refresh=refresh)
                        completed.add(item_id)
                        progress += 1
                        persist()
                    except Exception as exc:
                        failures.append({"id": item_id, "message": str(exc), "phase": "update"})
                    if max_messages and len(completed) - initial_count >= max_messages:
                        if on_progress:
                            on_progress(len(completed), -1)
                        if failures:
                            stopped("Some category updates failed.")
                        persist()
                        return len(completed)
                if on_progress:
                    on_progress(len(completed), -1)
                if failures:
                    stopped("Some category updates failed.")
                if not progress:
                    failures.append({"id": None, "message": "Repeated page without progress", "phase": "list"})
                    stopped("Category propagation made no progress.")
            stopped("Category propagation reached the 1000-page safety limit.")


def recolor_category(token: str, name: str, color: int) -> dict:
    """Change a master category's color."""
    return _update_master_categories(
        token,
        change_color=[{"Name": name, "Color": color}],
    )
