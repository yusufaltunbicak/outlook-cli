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
import uuid
from datetime import datetime, timezone
from typing import Callable
from urllib.parse import quote

import httpx

from .constants import BASE_URL, USER_AGENT

OWA_SERVICE_BASE = "https://outlook.cloud.microsoft/owa/service.svc"


def _owa_request(token: str, action: str, payload: dict) -> dict:
    """Send OWA service.svc request with x-owa-urlpostdata pattern."""
    resp = httpx.post(
        f"{OWA_SERVICE_BASE}?action={action}",
        headers={
            "Authorization": f"Bearer {token}",
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
        raise RuntimeError("Token expired. Run: outlook login")
    resp.raise_for_status()
    return resp.json()


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
        raise ValueError(f"Category '{old_name}' not found.")

    new_cat = {
        **existing,
        "Name": new_name,
        "LastTimeUsed": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
    }
    _update_master_categories(token, add=[new_cat], remove=[old_name])

    if not propagate:
        return 0

    return _bulk_rename_on_messages(token, old_name, new_name, on_progress)


def _bulk_rename_on_messages(
    token: str,
    old_name: str,
    new_name: str,
    on_progress: Callable[[int, int], None] | None = None,
) -> int:
    """Rename a category label on all messages that have it."""
    client = httpx.Client(
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30,
    )
    total = 0
    try:
        while True:
            r = client.get(
                f"{BASE_URL}/messages",
                params={
                    "$top": 50,
                    "$filter": f"Categories/any(c:c eq '{old_name}')",
                    "$select": "Id,Categories",
                },
            )
            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 5)))
                continue
            if r.status_code != 200:
                break
            msgs = r.json().get("value", [])
            if not msgs:
                break
            for m in msgs:
                new_cats = [new_name if c == old_name else c for c in m["Categories"]]
                for attempt in range(3):
                    try:
                        r2 = client.patch(
                            f"{BASE_URL}/messages/{m['Id']}",
                            json={"Categories": new_cats},
                        )
                        if r2.status_code == 200:
                            total += 1
                            break
                        elif r2.status_code == 429:
                            time.sleep(int(r2.headers.get("Retry-After", 5)))
                        else:
                            time.sleep(2)
                    except httpx.ReadTimeout:
                        time.sleep(3)
            if on_progress:
                on_progress(total, -1)
    finally:
        client.close()
    return total


def clear_category(
    token: str,
    name: str,
    folder: str | None = None,
    max_messages: int | None = None,
    on_progress: Callable[[int, int], None] | None = None,
) -> int:
    """Remove a category label from messages. Does not touch master category list.

    Args:
        folder: Limit to a specific folder (e.g. "Inbox"). None = all folders.
        max_messages: Stop after clearing this many messages. None = all.

    Returns the number of messages updated.
    """
    client = httpx.Client(
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30,
    )
    total = 0
    try:
        while True:
            if folder:
                url = f"{BASE_URL}/MailFolders/{folder}/messages"
            else:
                url = f"{BASE_URL}/messages"
            r = client.get(
                url,
                params={
                    "$top": 50,
                    "$filter": f"Categories/any(c:c eq '{name}')",
                    "$select": "Id,Categories",
                },
            )
            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 5)))
                continue
            if r.status_code != 200:
                break
            msgs = r.json().get("value", [])
            if not msgs:
                break
            for m in msgs:
                new_cats = [c for c in m["Categories"] if c != name]
                for attempt in range(3):
                    try:
                        r2 = client.patch(
                            f"{BASE_URL}/messages/{m['Id']}",
                            json={"Categories": new_cats},
                        )
                        if r2.status_code == 200:
                            total += 1
                            break
                        elif r2.status_code == 429:
                            time.sleep(int(r2.headers.get("Retry-After", 5)))
                        else:
                            time.sleep(2)
                    except httpx.ReadTimeout:
                        time.sleep(3)
                if max_messages and total >= max_messages:
                    break
            if on_progress:
                on_progress(total, -1)
            if max_messages and total >= max_messages:
                break
    finally:
        client.close()
    return total


def recolor_category(token: str, name: str, color: int) -> dict:
    """Change a master category's color."""
    return _update_master_categories(
        token,
        change_color=[{"Name": name, "Color": color}],
    )
