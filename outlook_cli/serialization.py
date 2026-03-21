from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime

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


def to_json_envelope(items: list | dict, pretty: bool = True, tz=None) -> str:
    """Wrap data in {ok, schema_version, data} envelope for stdout.

    When tz is provided, datetime values are converted to that timezone.
    """
    envelope = {
        "ok": True,
        "schema_version": SCHEMA_VERSION,
        "data": _normalize(items),
    }
    return json.dumps(envelope, cls=_encoder_cls(tz), indent=2 if pretty else None, ensure_ascii=False)


def error_json(code: str, message: str) -> str:
    """Structured error envelope for --json mode."""
    envelope = {
        "ok": False,
        "schema_version": SCHEMA_VERSION,
        "error": {"code": code, "message": message},
    }
    return json.dumps(envelope, indent=2, ensure_ascii=False)


def save_json(items: list | dict, path: str, tz=None) -> None:
    """Save raw JSON to file (no envelope — file export is raw data)."""
    with open(path, "w") as f:
        f.write(json.dumps(_normalize(items), cls=_encoder_cls(tz), indent=2, ensure_ascii=False))
