"""Tests for models.py — from_api() parsing from Outlook REST v2 JSON."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from outlook_cli.models import Attachment, Contact, Email, EmailAddress, Event, Folder

FIXTURES = Path(__file__).parent / "fixtures"


def _load(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text())


# ── Email ──────────────────────────────────────────────────


class TestEmailFromApi:
    def test_full_fields(self):
        data = _load("email_response.json")
        email = Email.from_api(data)

        assert email.id == data["Id"]
        assert email.subject == "Q4 Budget Review"
        assert email.sender.name == "Alice Smith"
        assert email.sender.address == "alice.smith@contoso.com"
        assert len(email.to) == 2
        assert email.to[0].address == "bob.jones@contoso.com"
        assert len(email.cc) == 1
        assert email.cc[0].name == "Dave Kim"
        assert email.is_read is False
        assert email.has_attachments is True
        assert email.importance == "High"
        assert email.categories == ["Finance", "Urgent"]
        assert email.body_type == "HTML"
        assert "budget proposal" in email.preview

    def test_scheduled_send_parsed(self):
        data = _load("email_response.json")
        email = Email.from_api(data)

        assert email.scheduled_send is not None
        assert email.scheduled_send.year == 2026
        assert email.scheduled_send.month == 3
        assert email.scheduled_send.day == 15

    def test_minimal_fields(self):
        """API returns only Id — everything else should default gracefully."""
        data = {"Id": "AAMk_minimal"}
        email = Email.from_api(data)

        assert email.id == "AAMk_minimal"
        assert email.subject == "(No Subject)"
        assert email.sender.address == ""
        assert email.to == []
        assert email.cc == []
        assert email.is_read is False
        assert email.has_attachments is False
        assert email.categories == []
        assert email.scheduled_send is None

    def test_no_scheduled_send(self):
        data = {"Id": "AAMk_nosched", "Subject": "Plain email"}
        email = Email.from_api(data)
        assert email.scheduled_send is None


# ── Event ──────────────────────────────────────────────────


class TestEventFromApi:
    def test_full_fields(self):
        data = _load("event_response.json")
        ev = Event.from_api(data)

        assert ev.subject == "Weekly Standup"
        assert ev.location == "Room A"
        assert ev.organizer.name == "Alice Smith"
        assert len(ev.attendees) == 2
        assert ev.attendees[0].response == "Accepted"
        assert ev.attendees[1].type == "Optional"
        assert ev.is_online_meeting is True
        assert "teams.microsoft.com" in ev.online_meeting_url
        assert ev.event_type == "SeriesMaster"
        assert ev.recurrence is not None
        assert ev.recurrence["Pattern"]["Type"] == "Weekly"

    def test_minimal_event(self):
        data = {
            "Id": "AAMk_evt_min",
            "Start": {"DateTime": "2026-03-16T10:00:00"},
            "End": {"DateTime": "2026-03-16T11:00:00"},
        }
        ev = Event.from_api(data)

        assert ev.id == "AAMk_evt_min"
        assert ev.subject == "(No Subject)"
        assert ev.attendees == []
        assert ev.recurrence is None
        assert ev.is_online_meeting is False


# ── Folder ─────────────────────────────────────────────────


class TestFolderFromApi:
    def test_parse(self):
        data = {
            "Id": "folder_inbox_123",
            "DisplayName": "Inbox",
            "UnreadItemCount": 12,
            "TotalItemCount": 350,
            "ParentFolderId": "root",
        }
        f = Folder.from_api(data)
        assert f.name == "Inbox"
        assert f.unread_count == 12
        assert f.total_count == 350


# ── Attachment ─────────────────────────────────────────────


class TestAttachmentFromApi:
    def test_parse(self):
        data = {
            "Id": "att_001",
            "Name": "report.pdf",
            "ContentType": "application/pdf",
            "Size": 204800,
            "IsInline": False,
            "ContentBytes": "JVBERi0xLjQ=",
        }
        att = Attachment.from_api(data)
        assert att.name == "report.pdf"
        assert att.size == 204800
        assert att.content_bytes == "JVBERi0xLjQ="


# ── Contact ────────────────────────────────────────────────


class TestContactFromApi:
    def test_parse(self):
        data = {
            "Id": "contact_001",
            "DisplayName": "Alice Smith",
            "GivenName": "Alice",
            "Surname": "Smith",
            "EmailAddresses": [
                {"Name": "Work", "Address": "alice@contoso.com"},
                {"Name": "Personal", "Address": "alice@gmail.com"},
            ],
            "CompanyName": "Contoso Ltd",
            "JobTitle": "CFO",
        }
        c = Contact.from_api(data)
        assert c.display_name == "Alice Smith"
        assert len(c.email_addresses) == 2
        assert c.company == "Contoso Ltd"


# ── EmailAddress ───────────────────────────────────────────


class TestEmailAddress:
    def test_str_with_name(self):
        ea = EmailAddress(name="Alice", address="alice@x.com")
        assert str(ea) == "Alice <alice@x.com>"

    def test_str_without_name(self):
        ea = EmailAddress(name="", address="alice@x.com")
        assert str(ea) == "alice@x.com"

    def test_from_api_nested(self):
        """API wraps in EmailAddress key."""
        data = {"EmailAddress": {"Name": "Bob", "Address": "bob@x.com"}}
        ea = EmailAddress.from_api(data)
        assert ea.name == "Bob"

    def test_from_api_flat(self):
        """Direct dict without EmailAddress wrapper."""
        data = {"Name": "Carol", "Address": "carol@x.com"}
        ea = EmailAddress.from_api(data)
        assert ea.name == "Carol"
