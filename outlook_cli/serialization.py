from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime

from .models import Attachment, Contact, Email, Event, Folder


class _Encoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, datetime):
            return o.isoformat()
        return super().default(o)


def to_json(items: list | dict, pretty: bool = True) -> str:
    if isinstance(items, list):
        data = [asdict(i) if hasattr(i, "__dataclass_fields__") else i for i in items]
    elif hasattr(items, "__dataclass_fields__"):
        data = asdict(items)
    else:
        data = items
    return json.dumps(data, cls=_Encoder, indent=2 if pretty else None, ensure_ascii=False)


def save_json(items: list | dict, path: str) -> None:
    with open(path, "w") as f:
        f.write(to_json(items))
