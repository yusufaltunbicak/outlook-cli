"""Cursor pagination with bounded work and explicit completeness metadata."""
from __future__ import annotations
from datetime import datetime, timezone
from urllib.parse import urljoin, urlsplit

from .constants import BASE_URL
from .exceptions import OutlookCliError


class Page(list):
    def __init__(self, values=(), *, meta=None):
        super().__init__(values)
        self.meta = dict(meta or {})


def validated_next_link(link: str, current_url: str, base_url: str = BASE_URL) -> str:
    if not isinstance(link, str):
        raise OutlookCliError("Invalid pagination nextLink")
    target = urljoin(current_url, link)
    base, candidate = urlsplit(base_url), urlsplit(target)
    if (candidate.scheme, candidate.hostname, candidate.port) != (base.scheme, base.hostname, base.port):
        raise OutlookCliError("Refusing pagination nextLink outside the API origin")
    if candidate.username or candidate.password or candidate.fragment:
        raise OutlookCliError("Refusing unsafe pagination nextLink")
    prefix = base.path.rstrip("/")
    if candidate.path != prefix and not candidate.path.startswith(prefix + "/"):
        raise OutlookCliError("Refusing pagination nextLink outside the mailbox API")
    return target


def paginate(get, path: str, params: dict, *, limit: int | None, predicate=None,
             search: bool = False, max_pages: int = 100, base_url: str = BASE_URL) -> Page:
    if "$top" in params and int(params["$top"]) < 1:
        raise ValueError("max must be at least 1")
    if limit is not None and limit < 1:
        raise ValueError("max must be at least 1")
    if max_pages < 1:
        raise ValueError("max_pages must be at least 1")
    rows, seen, cursors = [], set(), set()
    meta = {"complete": False, "has_more": False, "truncated_reason": None,
            "scope": "provider_search" if search else "provider_collection",
            "fetched_at": datetime.now(timezone.utc).isoformat(),
            "pages": 0, "fetched_count": 0, "returned_count": 0,
            "deduplicated_count": 0, "filtered_count": 0}
    request_path, request_params = path, params
    current_url = path if path.startswith("https://") else base_url.rstrip("/") + "/" + path.lstrip("/")
    for _ in range(max_pages):
        response = get(request_path, params=request_params)
        batch = response.get("value")
        if not isinstance(batch, list):
            raise OutlookCliError("Invalid API page: value must be a list")
        meta["pages"] += 1
        meta["fetched_count"] += len(batch)
        next_link = response.get("@odata.nextLink") or response.get("odata.nextLink")
        next_url = validated_next_link(next_link, current_url, base_url) if next_link else None
        unique = 0
        excess = False
        for item in batch:
            identity = item.get("Id") or item.get("id")
            if identity and identity in seen:
                meta["deduplicated_count"] += 1
                continue
            if identity:
                seen.add(identity)
            unique += 1
            if predicate is not None and not predicate(item):
                meta["filtered_count"] += 1
                continue
            if limit is not None and len(rows) >= limit:
                excess = True
                continue
            rows.append(item)
            if len(rows) >= 100000:
                excess = True
                break
        if (limit is not None and len(rows) >= limit) or len(rows) >= 100000:
            meta["has_more"] = bool(next_url or excess)
            meta["complete"] = not meta["has_more"] and not search
            meta["truncated_reason"] = "limit" if meta["has_more"] else ("search_scope_unknown" if search else None)
            if search and not meta["has_more"]:
                meta["has_more"] = None
            break
        if not next_url:
            meta["complete"] = not search
            meta["has_more"] = None if search else False
            meta["truncated_reason"] = "search_scope_unknown" if search else None
            break
        meta["has_more"] = True
        if next_url in cursors:
            meta["truncated_reason"] = "repeated_next_link"
            break
        if batch and unique == 0:
            meta["truncated_reason"] = "repeated_page"
            break
        cursors.add(next_url)
        request_path, request_params, current_url = next_url, None, next_url
    else:
        meta["truncated_reason"] = "page_limit"
        meta["has_more"] = True
    meta["returned_count"] = len(rows)
    return Page(rows, meta=meta)
