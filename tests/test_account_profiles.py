"""Tests for account registry, precedence, and profile-scoped storage."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from outlook_cli import account as account_service
from outlook_cli import auth as auth_mod
from outlook_cli import client as client_mod
from outlook_cli import signature_manager as signature_manager_mod
from outlook_cli.account import AccountPaths
from outlook_cli.client import OutlookClient
from outlook_cli.exceptions import AccountError


def _patch_account_roots(monkeypatch, tmp_path: Path):
    cache_root = tmp_path / "cache"
    config_root = tmp_path / "config"
    monkeypatch.setattr(account_service, "CACHE_DIR", cache_root)
    monkeypatch.setattr(account_service, "CONFIG_DIR", config_root)
    monkeypatch.setattr(account_service, "TOKEN_FILE", cache_root / "token.json")
    monkeypatch.setattr(account_service, "BROWSER_STATE_FILE", cache_root / "browser-state.json")
    monkeypatch.setattr(account_service, "ID_MAP_FILE", cache_root / "id_map.json")
    monkeypatch.setattr(account_service, "SCHEDULED_FILE", cache_root / "scheduled.json")
    monkeypatch.setattr(account_service, "SIGNATURES_DIR", config_root / "signatures")
    monkeypatch.setattr(account_service, "CONFIG_FILE", config_root / "config.yaml")
    monkeypatch.setattr(account_service, "ACCOUNTS_FILE", config_root / "accounts.json")
    monkeypatch.setattr(account_service, "ACCOUNTS_CACHE_DIR", cache_root / "accounts")
    monkeypatch.setattr(account_service, "ACCOUNTS_CONFIG_DIR", config_root / "accounts")
    return cache_root, config_root


def _paths_for(root: Path, name: str) -> AccountPaths:
    return AccountPaths(
        name=name,
        cache_dir=root / "cache" / "accounts" / name,
        config_dir=root / "config" / "accounts" / name,
        token_file=root / "cache" / "accounts" / name / "token.json",
        browser_state_file=root / "cache" / "accounts" / name / "browser-state.json",
        id_map_file=root / "cache" / "accounts" / name / "id_map.json",
        scheduled_file=root / "cache" / "accounts" / name / "scheduled.json",
        signatures_dir=root / "config" / "accounts" / name / "signatures",
        profile_config_file=root / "config" / "accounts" / name / "config.yaml",
    )


def test_resolve_account_name_uses_precedence(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps(
            {
                "current_account": "work",
                "accounts": {
                    "work": {"name": "work"},
                    "personal": {"name": "personal"},
                },
            }
        )
    )

    assert account_service.resolve_account_name() == "work"
    monkeypatch.setenv("OUTLOOK_ACCOUNT", "personal")
    assert account_service.resolve_account_name() == "personal"
    assert account_service.resolve_account_name("work") == "work"


def test_default_account_uses_legacy_paths_until_profile_dirs_exist(monkeypatch, tmp_path):
    cache_root, config_root = _patch_account_roots(monkeypatch, tmp_path)

    legacy = account_service.get_account_paths("default")
    assert legacy.uses_legacy_default is True
    assert legacy.token_file == cache_root / "token.json"

    (config_root / "accounts" / "default").mkdir(parents=True, exist_ok=True)
    profile = account_service.get_account_paths("default")
    assert profile.uses_legacy_default is False
    assert profile.token_file == cache_root / "accounts" / "default" / "token.json"


def test_bind_account_rejects_duplicate_mailbox(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps(
            {
                "current_account": "work",
                "accounts": {
                    "work": {"name": "work", "mailbox_id": "mailbox-1", "email": "user@example.com"},
                },
            }
        )
    )

    with pytest.raises(AccountError, match="already bound"):
        account_service.bind_account(
            "personal",
            {"Id": "mailbox-1", "EmailAddress": "user@example.com", "DisplayName": "User"},
        )


def test_remove_account_deletes_stored_keyring_token(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps(
            {
                "current_account": "default",
                "accounts": {
                    "work": {"name": "work", "mailbox_id": "mailbox-1", "email": "work@example.com"},
                },
            }
        )
    )
    paths = account_service.get_account_paths("work")
    paths.cache_dir.mkdir(parents=True, exist_ok=True)
    paths.config_dir.mkdir(parents=True, exist_ok=True)
    deleted = []
    monkeypatch.setattr(auth_mod, "delete_stored_token", lambda name=None: deleted.append(name))

    account_service.remove_account("work")

    assert deleted == ["work"]


def test_load_account_config_merges_global_and_profile_overrides(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.CONFIG_FILE.write_text("max_messages: 50\nbrowser:\n  timeout: 300\ntimezone: UTC\n")
    profile_dir = account_service.ACCOUNTS_CONFIG_DIR / "work"
    profile_dir.mkdir(parents=True, exist_ok=True)
    (profile_dir / "config.yaml").write_text("timezone: Europe/Istanbul\ndefault_signature: work\n")

    cfg = account_service.load_account_config("work")

    assert cfg["max_messages"] == 50
    assert cfg["browser"]["timeout"] == 300
    assert cfg["timezone"] == "Europe/Istanbul"
    assert cfg["default_signature"] == "work"


def test_profile_scoped_signatures_are_isolated(monkeypatch, tmp_path):
    root = tmp_path
    path_map = {
        "work": _paths_for(root, "work"),
        "personal": _paths_for(root, "personal"),
    }
    monkeypatch.setattr(
        signature_manager_mod.account_service,
        "resolve_account_name",
        lambda account_name=None: account_name or "work",
    )
    monkeypatch.setattr(signature_manager_mod.account_service, "get_account_paths", lambda name: path_map[name])

    signature_manager_mod.save_signature("default", "<b>Work</b>", account_name="work")
    signature_manager_mod.save_signature("default", "<b>Personal</b>", account_name="personal")

    assert signature_manager_mod.get_signature("default", account_name="work") == "<b>Work</b>"
    assert signature_manager_mod.get_signature("default", account_name="personal") == "<b>Personal</b>"


def test_profile_scoped_id_maps_and_schedules_are_isolated(monkeypatch, tmp_path):
    root = tmp_path
    path_map = {
        "work": _paths_for(root, "work"),
        "personal": _paths_for(root, "personal"),
    }
    monkeypatch.setattr(
        client_mod.account_service,
        "resolve_account_name",
        lambda account_name=None, allow_missing=False: account_name or "work",
    )
    monkeypatch.setattr(client_mod.account_service, "get_account_paths", lambda name: path_map[name])

    work = OutlookClient("token", account_name="work")
    personal = OutlookClient("token", account_name="personal")

    work._id_map = {"1": "work-id"}
    work._save_id_map()
    personal._id_map = {"1": "personal-id"}
    personal._save_id_map()

    work._save_scheduled([{"subject": "Work", "scheduled_at": "2026-03-17T10:00:00Z"}])
    personal._save_scheduled([{"subject": "Personal", "scheduled_at": "2026-03-17T11:00:00Z"}])

    assert json.loads(path_map["work"].id_map_file.read_text()) == {"1": "work-id"}
    assert json.loads(path_map["personal"].id_map_file.read_text()) == {"1": "personal-id"}
    assert work._load_scheduled()[0]["subject"] == "Work"
    assert personal._load_scheduled()[0]["subject"] == "Personal"
