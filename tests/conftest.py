from __future__ import annotations

from datetime import datetime, timezone

import pytest
from click.testing import CliRunner

from outlook_cli.commands import _common as common
from outlook_cli.models import Attachment, Contact, Email, EmailAddress, Event, Folder


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
