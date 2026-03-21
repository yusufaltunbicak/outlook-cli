"""Tests for serialization.py — envelope format, error_json, encoding."""

from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone

from outlook_cli.models import Email, EmailAddress, Folder
from outlook_cli.serialization import error_json, save_json, to_json, to_json_envelope


class TestToJsonEnvelope:
    def test_success_envelope_structure(self):
        result = json.loads(to_json_envelope({"key": "value"}))
        assert result["ok"] is True
        assert result["schema_version"] == "1"
        assert result["data"] == {"key": "value"}

    def test_envelope_with_list(self):
        result = json.loads(to_json_envelope([1, 2, 3]))
        assert result["ok"] is True
        assert result["data"] == [1, 2, 3]

    def test_envelope_with_dataclass(self):
        folder = Folder(id="f1", name="Inbox", unread_count=5, total_count=100, parent_folder_id="root")
        result = json.loads(to_json_envelope(folder))
        assert result["ok"] is True
        assert result["data"]["name"] == "Inbox"
        assert result["data"]["unread_count"] == 5

    def test_envelope_with_dataclass_list(self):
        folders = [
            Folder(id="f1", name="Inbox", unread_count=5, total_count=100, parent_folder_id="root"),
            Folder(id="f2", name="Sent", unread_count=0, total_count=50, parent_folder_id="root"),
        ]
        result = json.loads(to_json_envelope(folders))
        assert result["ok"] is True
        assert len(result["data"]) == 2
        assert result["data"][1]["name"] == "Sent"

    def test_envelope_normalizes_nested_dataclasses(self):
        email = Email(
            id="m1",
            subject="Hello",
            sender=EmailAddress(name="Alice", address="alice@example.com"),
            to=[],
            cc=[],
            received=datetime(2026, 3, 15, 10, 0, tzinfo=timezone.utc),
            preview="",
            body="",
            body_type="Text",
            is_read=True,
            has_attachments=False,
            importance="Normal",
            conversation_id="conv-1",
        )

        result = json.loads(to_json_envelope({"messages": [email]}))
        assert result["data"]["messages"][0]["subject"] == "Hello"


class TestErrorJson:
    def test_error_envelope_structure(self):
        result = json.loads(error_json("not_found", "Message #999 not found"))
        assert result["ok"] is False
        assert result["schema_version"] == "1"
        assert result["error"]["code"] == "not_found"
        assert result["error"]["message"] == "Message #999 not found"

    def test_error_codes(self):
        for code in ["session_expired", "rate_limited", "not_found", "not_authenticated", "unknown_error"]:
            result = json.loads(error_json(code, "test"))
            assert result["error"]["code"] == code


class TestToJsonRaw:
    def test_no_envelope(self):
        """to_json() should NOT include envelope — it's for file export."""
        result = json.loads(to_json({"subject": "Hello"}))
        assert "ok" not in result
        assert result["subject"] == "Hello"

    def test_datetime_encoding(self):
        dt = datetime(2026, 3, 15, 10, 0, 0, tzinfo=timezone.utc)
        result = json.loads(to_json({"when": dt}))
        assert "2026-03-15" in result["when"]


class TestSaveJson:
    def test_writes_raw_to_file(self):
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="r") as f:
            path = f.name
        save_json({"test": True}, path)
        data = json.loads(open(path).read())
        assert "ok" not in data  # no envelope
        assert data["test"] is True
