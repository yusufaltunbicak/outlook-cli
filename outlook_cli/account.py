from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import yaml

from .config import _deep_merge, load_config
from .constants import ACCOUNTS_CACHE_DIR, ACCOUNTS_CONFIG_DIR, ACCOUNTS_FILE, BROWSER_STATE_FILE, CACHE_DIR, CONFIG_DIR, CONFIG_FILE, ID_MAP_FILE, SCHEDULED_FILE, SIGNATURES_DIR, TOKEN_FILE
from .exceptions import AccountError

ACCOUNT_NAME_RE = re.compile(r"^[a-z0-9][a-z0-9_-]*$")


@dataclass(frozen=True)
class AccountPaths:
    name: str
    cache_dir: Path
    config_dir: Path
    token_file: Path
    browser_state_file: Path
    id_map_file: Path
    scheduled_file: Path
    signatures_dir: Path
    profile_config_file: Path
    uses_legacy_default: bool = False


def normalize_account_name(name: str) -> str:
    normalized = name.strip().lower()
    if not normalized or not ACCOUNT_NAME_RE.match(normalized):
        raise AccountError("Account profile names must match [a-z0-9][a-z0-9_-]*.")
    return normalized


def _empty_registry() -> dict[str, Any]:
    return {"current_account": None, "accounts": {}}


def load_registry() -> dict[str, Any]:
    if not ACCOUNTS_FILE.exists():
        return _empty_registry()

    try:
        data = json.loads(ACCOUNTS_FILE.read_text())
    except (json.JSONDecodeError, OSError):
        return _empty_registry()

    accounts = data.get("accounts")
    if not isinstance(accounts, dict):
        accounts = {}

    current = data.get("current_account")
    if current is not None:
        try:
            current = normalize_account_name(current)
        except AccountError:
            current = None

    cleaned: dict[str, dict[str, Any]] = {}
    for raw_name, meta in accounts.items():
        try:
            name = normalize_account_name(raw_name)
        except AccountError:
            continue
        cleaned[name] = {
            "name": name,
            "mailbox_id": meta.get("mailbox_id"),
            "email": meta.get("email"),
            "display_name": meta.get("display_name"),
            "created_at": meta.get("created_at"),
            "last_used_at": meta.get("last_used_at"),
            "legacy_default": bool(meta.get("legacy_default", False)),
        }

    return {"current_account": current, "accounts": cleaned}


def save_registry(registry: dict[str, Any]) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "current_account": registry.get("current_account"),
        "accounts": registry.get("accounts", {}),
    }
    ACCOUNTS_FILE.write_text(json.dumps(payload, indent=2))


def resolve_account_name(explicit_name: str | None = None, *, allow_missing: bool = False) -> str:
    name = explicit_name or os.environ.get("OUTLOOK_ACCOUNT")
    if name:
        resolved = normalize_account_name(name)
    else:
        registry = load_registry()
        current = registry.get("current_account")
        resolved = normalize_account_name(current) if current else "default"

    if not allow_missing:
        ensure_account_known(resolved)

    return resolved


def ensure_account_known(name: str, registry: dict[str, Any] | None = None) -> None:
    registry = registry or load_registry()
    if name == "default":
        return
    if name not in registry.get("accounts", {}):
        raise AccountError(
            f"Account profile '{name}' not found. Run 'outlook account add {name}' first."
        )


def get_current_account_name() -> str:
    registry = load_registry()
    current = registry.get("current_account")
    return normalize_account_name(current) if current else "default"


def set_current_account(name: str) -> None:
    name = normalize_account_name(name)
    ensure_account_known(name)
    registry = load_registry()
    registry["current_account"] = name
    save_registry(registry)


def _profile_cache_dir(name: str) -> Path:
    return ACCOUNTS_CACHE_DIR / name


def _profile_config_dir(name: str) -> Path:
    return ACCOUNTS_CONFIG_DIR / name


def uses_legacy_default_paths(name: str) -> bool:
    return name == "default" and not _profile_cache_dir("default").exists() and not _profile_config_dir("default").exists()


def get_account_paths(name: str) -> AccountPaths:
    name = normalize_account_name(name)
    cache_dir = _profile_cache_dir(name)
    config_dir = _profile_config_dir(name)
    profile_config_file = config_dir / "config.yaml"
    if uses_legacy_default_paths(name):
        return AccountPaths(
            name=name,
            cache_dir=CACHE_DIR,
            config_dir=CONFIG_DIR,
            token_file=TOKEN_FILE,
            browser_state_file=BROWSER_STATE_FILE,
            id_map_file=ID_MAP_FILE,
            scheduled_file=SCHEDULED_FILE,
            signatures_dir=SIGNATURES_DIR,
            profile_config_file=profile_config_file,
            uses_legacy_default=True,
        )
    return AccountPaths(
        name=name,
        cache_dir=cache_dir,
        config_dir=config_dir,
        token_file=cache_dir / "token.json",
        browser_state_file=cache_dir / "browser-state.json",
        id_map_file=cache_dir / "id_map.json",
        scheduled_file=cache_dir / "scheduled.json",
        signatures_dir=config_dir / "signatures",
        profile_config_file=profile_config_file,
        uses_legacy_default=False,
    )


def has_legacy_default_state() -> bool:
    return any(
        path.exists()
        for path in (TOKEN_FILE, BROWSER_STATE_FILE, ID_MAP_FILE, SCHEDULED_FILE, SIGNATURES_DIR)
    )


def mailbox_info_from_me(me: dict[str, Any]) -> dict[str, str]:
    email = (me.get("EmailAddress") or me.get("email") or "").strip()
    mailbox_id = str(me.get("Id") or me.get("mailbox_id") or email.lower())
    if not mailbox_id:
        raise AccountError("Could not determine mailbox identity for the authenticated account.")
    return {
        "mailbox_id": mailbox_id,
        "email": email,
        "display_name": (me.get("DisplayName") or me.get("display_name") or "").strip(),
    }


def _same_mailbox(left: dict[str, Any], right: dict[str, Any]) -> bool:
    left_id = (left.get("mailbox_id") or "").strip().lower()
    right_id = (right.get("mailbox_id") or "").strip().lower()
    left_email = (left.get("email") or "").strip().lower()
    right_email = (right.get("email") or "").strip().lower()
    if left_id and right_id and left_id == right_id:
        return True
    return bool(left_email and right_email and left_email == right_email)


def get_account(name: str, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    name = normalize_account_name(name)
    registry = registry or load_registry()
    meta = dict(registry.get("accounts", {}).get(name, {}))
    if not meta:
        meta = {"name": name}
    meta.setdefault("name", name)
    if name == "default":
        meta.setdefault("legacy_default", uses_legacy_default_paths(name))
    return meta


def bind_account(name: str, me: dict[str, Any]) -> dict[str, Any]:
    name = normalize_account_name(name)
    info = mailbox_info_from_me(me)
    registry = load_registry()
    for existing_name, meta in registry.get("accounts", {}).items():
        if existing_name != name and _same_mailbox(meta, info):
            other = meta.get("email") or meta.get("display_name") or existing_name
            raise AccountError(
                f"Mailbox '{other}' is already bound to account profile '{existing_name}'."
            )

    existing = registry.get("accounts", {}).get(name, {})
    now = datetime.now(timezone.utc).isoformat()
    merged = {
        "name": name,
        "mailbox_id": info["mailbox_id"],
        "email": info["email"],
        "display_name": info["display_name"],
        "created_at": existing.get("created_at") or now,
        "last_used_at": now,
        "legacy_default": name == "default" and uses_legacy_default_paths(name),
    }
    registry.setdefault("accounts", {})[name] = merged
    save_registry(registry)
    return merged


def assert_mailbox_matches(name: str, me: dict[str, Any]) -> dict[str, Any]:
    name = normalize_account_name(name)
    info = mailbox_info_from_me(me)
    bound = get_account(name)
    if bound.get("mailbox_id") and not _same_mailbox(bound, info):
        expected = bound.get("email") or bound.get("display_name") or name
        actual = info.get("email") or info.get("display_name") or info.get("mailbox_id")
        raise AccountError(
            f"Authenticated mailbox '{actual}' does not match account profile '{name}' ({expected})."
        )
    return info


def touch_account(name: str) -> None:
    name = normalize_account_name(name)
    registry = load_registry()
    meta = registry.get("accounts", {}).get(name)
    if not meta:
        return
    meta["last_used_at"] = datetime.now(timezone.utc).isoformat()
    registry["accounts"][name] = meta
    save_registry(registry)


def list_accounts() -> list[dict[str, Any]]:
    registry = load_registry()
    current = registry.get("current_account") or "default"
    names = set(registry.get("accounts", {}))
    names.add("default")
    rows = []
    for name in sorted(names, key=lambda value: (value != "default", value)):
        meta = get_account(name, registry)
        rows.append(
            {
                "name": name,
                "current": current == name,
                "bound": bool(meta.get("mailbox_id")),
                "mailbox_id": meta.get("mailbox_id"),
                "email": meta.get("email"),
                "display_name": meta.get("display_name"),
                "created_at": meta.get("created_at"),
                "last_used_at": meta.get("last_used_at"),
                "legacy_default": bool(meta.get("legacy_default", False)),
            }
        )
    return rows


def remove_account(name: str) -> None:
    name = normalize_account_name(name)
    current = get_current_account_name()
    if name == current:
        raise AccountError(f"Cannot remove the current account profile '{name}'. Switch first.")

    registry = load_registry()
    if name != "default":
        ensure_account_known(name, registry)

    paths = get_account_paths(name)
    if paths.uses_legacy_default:
        for path in (TOKEN_FILE, BROWSER_STATE_FILE, ID_MAP_FILE, SCHEDULED_FILE):
            if path.exists():
                path.unlink()
        if SIGNATURES_DIR.exists():
            shutil.rmtree(SIGNATURES_DIR)
    else:
        if paths.cache_dir.exists():
            shutil.rmtree(paths.cache_dir)
        if paths.config_dir.exists():
            shutil.rmtree(paths.config_dir)

    registry.get("accounts", {}).pop(name, None)
    save_registry(registry)


def load_account_config(name: str) -> dict[str, Any]:
    name = normalize_account_name(name)
    cfg = load_config(CONFIG_FILE)
    profile_config = get_account_paths(name).profile_config_file
    if profile_config.exists():
        with profile_config.open() as f:
            user_cfg = yaml.safe_load(f) or {}
        _deep_merge(cfg, user_cfg)
    return cfg


def current_account_snapshot() -> dict[str, Any]:
    name = get_current_account_name()
    meta = get_account(name)
    return {
        "name": name,
        "current": True,
        "bound": bool(meta.get("mailbox_id")),
        "mailbox_id": meta.get("mailbox_id"),
        "email": meta.get("email"),
        "display_name": meta.get("display_name"),
        "created_at": meta.get("created_at"),
        "last_used_at": meta.get("last_used_at"),
        "legacy_default": bool(meta.get("legacy_default", False)),
        "paths": asdict(get_account_paths(name)),
    }
