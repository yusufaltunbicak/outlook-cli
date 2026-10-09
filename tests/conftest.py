from __future__ import annotations

import socket
from datetime import datetime, timezone

import httpx
import keyring
import pytest
from click.testing import CliRunner

from outlook_cli import account, config, constants, credentials
from outlook_cli.commands import _common as common
from outlook_cli.models import Attachment, Contact, Email, EmailAddress, Event, Folder


@pytest.fixture(autouse=True)
def offline_unit_environment(monkeypatch, tmp_path, request):
    """Unit tests cannot touch a real mailbox, credential store or local cache.

    Tests replace the blocked boundary explicitly when a fake transport,
    browser or keyring is part of the scenario. Live smoke tests remain a
    separate, explicitly selected read-only suite.
    """
    if request.node.get_closest_marker("smoke"):
        return

    def denied(*_args, **_kwargs):
        raise AssertionError("Live network, browser and keychain access is forbidden in unit tests")

    cache_root = tmp_path / "isolated-cache"
    config_root = tmp_path / "isolated-config"
    paths = {
        "CACHE_DIR": cache_root,
        "CONFIG_DIR": config_root,
        "TOKEN_FILE": cache_root / "token.json",
        "BROWSER_STATE_FILE": cache_root / "browser-state.json",
        "ID_MAP_FILE": cache_root / "id_map.json",
        "SCHEDULED_FILE": cache_root / "scheduled.json",
        "SIGNATURES_DIR": config_root / "signatures",
        "CONFIG_FILE": config_root / "config.yaml",
        "ACCOUNTS_FILE": config_root / "accounts.json",
        "ACCOUNTS_CACHE_DIR": cache_root / "accounts",
        "ACCOUNTS_CONFIG_DIR": config_root / "accounts",
    }
    for module in (constants, account, config):
        for name, value in paths.items():
            if hasattr(module, name):
                monkeypatch.setattr(module, name, value)
    monkeypatch.setenv("OUTLOOK_CLI_CACHE", str(cache_root))
    monkeypatch.setenv("OUTLOOK_CLI_CONFIG", str(config_root))
    for name in ("OUTLOOK_TOKEN", "OUTLOOK_ACCOUNT", "OUTLOOK_GRAPH_TOKEN"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setattr(common, "_client_cache", {})
    monkeypatch.setattr(common.cfg, "_overrides", {})
    monkeypatch.setattr(socket.socket, "connect", denied)
    monkeypatch.setattr(socket.socket, "connect_ex", denied)
    monkeypatch.setattr(httpx.Client, "send", denied)
    monkeypatch.setattr(httpx.AsyncClient, "send", denied)
    for name in ("get_password", "set_password", "delete_password"):
        monkeypatch.setattr(keyring, name, denied)
    monkeypatch.setattr(credentials, "_native_call", denied)
    from playwright import sync_api
    monkeypatch.setattr(sync_api, "sync_playwright", denied)


class DummyResponse:
    def __init__(
        self,
        status_code: int = 200,
        json_data: dict | None = None,
        headers: dict | None = None,
        content: bytes | None = None,
    ):
        self.status_code = status_code
        self._json_data = json_data or {}
        self.headers = headers or {}
        if content is not None:
            self.content = content
        elif status_code == 204:
            self.content = b""
        else:
            self.content = b"{}"

    def json(self) -> dict:
        return self._json_data

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            import httpx

            request = httpx.Request("GET", "https://example.com")
            response = httpx.Response(self.status_code, request=request)
            raise httpx.HTTPStatusError("request failed", request=request, response=response)


@pytest.fixture
def runner() -> CliRunner:
    return CliRunner()


@pytest.fixture
def tty_mode(monkeypatch):
    monkeypatch.setattr(common, "_is_piped", lambda: False)
    monkeypatch.setattr(common, "_stdin_is_tty", lambda: True)


@pytest.fixture
def make_email():
    def _make_email(**overrides) -> Email:
        base = {
            "id": "msg-1",
            "subject": "Subject",
            "sender": EmailAddress(name="Alice", address="alice@example.com"),
            "to": [EmailAddress(name="Bob", address="bob@example.com")],
            "cc": [],
            "received": datetime(2026, 3, 17, 9, 0, tzinfo=timezone.utc),
            "preview": "Preview",
            "body": "Body",
            "body_type": "Text",
            "is_read": True,
            "has_attachments": False,
            "importance": "Normal",
            "conversation_id": "conv-1",
            "categories": [],
            "flag_status": "notFlagged",
            "flag_due": None,
            "scheduled_send": None,
            "display_num": 1,
        }
        base.update(overrides)
        return Email(**base)

    return _make_email


@pytest.fixture
def make_event():
    def _make_event(**overrides) -> Event:
        base = {
            "id": "ev-1",
            "subject": "Standup",
            "start": datetime(2026, 3, 17, 10, 0, tzinfo=timezone.utc),
            "end": datetime(2026, 3, 17, 11, 0, tzinfo=timezone.utc),
            "location": "Room A",
            "organizer": EmailAddress(name="Alice", address="alice@example.com"),
            "is_all_day": False,
            "body_preview": "Preview",
            "body": "Body",
            "body_type": "Text",
            "attendees": [],
            "categories": [],
            "show_as": "Busy",
            "sensitivity": "Normal",
            "is_cancelled": False,
            "response_status": "",
            "web_link": "",
            "is_online_meeting": False,
            "online_meeting_url": "",
            "recurrence": None,
            "event_type": "SingleInstance",
            "series_master_id": "",
            "display_num": 1,
        }
        base.update(overrides)
        return Event(**base)

    return _make_event


@pytest.fixture
def make_folder():
    def _make_folder(**overrides) -> Folder:
        base = {
            "id": "folder-1",
            "name": "Inbox",
            "unread_count": 2,
            "total_count": 10,
            "parent_folder_id": "root",
        }
        base.update(overrides)
        return Folder(**base)

    return _make_folder


@pytest.fixture
def make_contact():
    def _make_contact(**overrides) -> Contact:
        base = {
            "id": "contact-1",
            "display_name": "Alice Smith",
            "given_name": "Alice",
            "surname": "Smith",
            "email_addresses": [EmailAddress(name="Work", address="alice@example.com")],
            "company": "Contoso",
            "job_title": "CFO",
        }
        base.update(overrides)
        return Contact(**base)

    return _make_contact


@pytest.fixture
def make_attachment():
    def _make_attachment(**overrides) -> Attachment:
        base = {
            "id": "att-1",
            "name": "report.pdf",
            "content_type": "application/pdf",
            "size": 1024,
            "is_inline": False,
            "content_bytes": None,
        }
        base.update(overrides)
        return Attachment(**base)

    return _make_attachment
