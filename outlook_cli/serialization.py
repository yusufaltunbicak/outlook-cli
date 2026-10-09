from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

import click

from .models import Attachment, Contact, Email, Event, Folder

SCHEMA_VERSION = "1"


class _Encoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, datetime):
            return o.isoformat()
        return super().default(o)


def _encoder_cls(tz=None):
    """Create encoder class with optional timezone conversion."""
    if tz is None:
        return _Encoder

    class _TzEncoder(json.JSONEncoder):
        def default(self, o):
            if isinstance(o, datetime):
                if o.tzinfo:
                    return o.astimezone(tz).isoformat()
                return o.isoformat()
            return super().default(o)

    return _TzEncoder


def _normalize(items):
    """Convert dataclasses / mixed lists to plain dicts."""
    if hasattr(items, "__dataclass_fields__"):
        return _normalize(asdict(items))
    if isinstance(items, list):
        return [_normalize(i) for i in items]
    if isinstance(items, tuple):
        return [_normalize(i) for i in items]
    if isinstance(items, dict):
        return {key: _normalize(value) for key, value in items.items()}
    return items


def to_json(items: list | dict, pretty: bool = True) -> str:
    """Raw JSON — used by save_json for file export."""
    return json.dumps(_normalize(items), cls=_Encoder, indent=2 if pretty else None, ensure_ascii=False)


def output_settings() -> dict:
    """Get serialization preferences without coupling Python callers to Click."""
    ctx = click.get_current_context(silent=True)
    return dict(getattr(ctx, "_outlook_output", {})) if ctx else {}


def _project(data, settings):
    if isinstance(data, list):
        return [_project(item, settings) for item in data]
    if not isinstance(data, dict):
        return data
    data = dict(data)
    # Apply body presentation only to message-shaped objects.
    if "body" in data and ("sender" in data or "conversation_id" in data):
        mode = settings.get("body_format")
        if mode == "none":
            data.pop("body", None)
            data.pop("body_type", None)
        elif mode == "preview":
            data["body"] = data.get("preview", "")
            data["body_type"] = "Text"
        elif mode == "text" and str(data.get("body_type", "")).lower() == "html":
            from bs4 import BeautifulSoup
            data["body"] = BeautifulSoup(data["body"], "html.parser").get_text("\n", strip=True)
            data["body_type"] = "Text"
    if "ok" in data and ("data" in data or "error" in data):
        if "data" in data:
            data["data"] = _project(data["data"], settings)
        return data
    fields = settings.get("fields")
    if fields:
        return {key: data.get(key) for key in fields}
    if settings.get("view") == "compact" and "sender" in data:
        return {key: data[key] for key in COMPACT_FIELDS if key in data}
    # Nested bulk results retain their status/error structure.
    for key in ("data", "messages", "results"):
        if key in data:
            data[key] = _project(data[key], settings)
    return data


COMPACT_FIELDS = ("id", "display_num", "subject", "sender", "received", "preview", "is_read", "has_attachments", "categories", "conversation_id")
MAIL_API_FIELDS = {
    "id": "Id", "display_num": "Id", "subject": "Subject", "sender": "From",
    "to": "ToRecipients", "cc": "CcRecipients", "received": "ReceivedDateTime",
    "preview": "BodyPreview", "body": "Body", "body_type": "Body", "is_read": "IsRead",
    "has_attachments": "HasAttachments", "importance": "Importance", "conversation_id": "ConversationId",
    "categories": "Categories", "flag_status": "Flag", "flag_due": "Flag",
}


def message_fetch_options() -> dict:
    settings = output_settings()
    fields = settings.get("fields") or (COMPACT_FIELDS if settings.get("view") == "compact" else None)
    opts = {}
    if fields:
        opts["select"] = ",".join(dict.fromkeys(["Id"] + [MAIL_API_FIELDS[k] for k in fields if k in MAIL_API_FIELDS]))
    elif settings.get("body_format") in {"none", "preview"}:
        opts["select"] = ",".join(dict.fromkeys(v for k, v in MAIL_API_FIELDS.items() if v != "Body"))
    if settings.get("all_pages"):
        opts["all_pages"] = True
    return opts


def pagination_options() -> dict:
    return {"all_pages": True} if output_settings().get("all_pages") else {}


def _record_output(text: str) -> None:
    ctx = click.get_current_context(silent=True)
    if ctx is not None:
        ctx._outlook_emitted = True
        path = output_settings().get("output")
        if path:
            Path(path).write_text(text + "\n", encoding="utf-8")


def to_json_envelope(items: list | dict, pretty: bool = True, tz=None, *, meta=None, ok=True, error=None) -> str:
    """Canonical stdout/file contract; raw data is explicit via --data-only."""
    settings = output_settings()
    data = _project(_normalize(items), settings)
    envelope = {"ok": ok, "schema_version": SCHEMA_VERSION, "data": data}
    page_meta = getattr(items, "meta", None)
    if isinstance(page_meta, dict) or meta:
        envelope["meta"] = {**(page_meta if isinstance(page_meta, dict) else {}), **(meta or {})}
    if error:
        envelope["error"] = error
    result = data if settings.get("data_only") and ok else envelope
    from .telemetry import timed_phase
    with timed_phase("serialization"):
        text = json.dumps(result, cls=_encoder_cls(tz), indent=2 if pretty else None, ensure_ascii=False)
    _record_output(text)
    return text


def error_json(code: str, message: str) -> str:
    envelope = {"ok": False, "schema_version": SCHEMA_VERSION, "error": {"code": code, "message": message}}
    text = json.dumps(envelope, indent=2, ensure_ascii=False)
    _record_output(text)
    return text


def save_json(items: list | dict, path: str, tz=None) -> None:
    """CLI exports match stdout; Python API retains its historical raw export."""
    if click.get_current_context(silent=True) is not None:
        text = to_json_envelope(items, tz=tz)
        Path(path).write_text(text + "\n", encoding="utf-8")
        click.echo(text)
    else:
        Path(path).write_text(json.dumps(_normalize(items), cls=_encoder_cls(tz), indent=2, ensure_ascii=False), encoding="utf-8")
