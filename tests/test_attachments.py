"""Tests for file attachment support — client methods and command integration."""

from __future__ import annotations

import base64
import json
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

from outlook_cli.client import OutlookClient
from outlook_cli.constants import ATTACHMENT_SIZE_THRESHOLD


# ── Fixtures ──────────────────────────────────────────────


FAKE_MSG_ID = "AAMk_test_message_id_long_enough_to_pass_resolve_check_1234567890"


@pytest.fixture
def client():
    """Create an OutlookClient with mocked HTTP and no id_map file."""
    with patch.object(OutlookClient, "_load_id_map", return_value={}):
        c = OutlookClient("fake-token")
    return c


@pytest.fixture
def small_file(tmp_path):
    """Create a small text file (well under 3 MB)."""
    f = tmp_path / "small.txt"
    f.write_text("Hello, World!")
    return f


@pytest.fixture
def large_file(tmp_path):
    """Create a file larger than the 3 MB threshold."""
    f = tmp_path / "large.bin"
    f.write_bytes(b"\x00" * (ATTACHMENT_SIZE_THRESHOLD + 1024))
    return f


# ── add_attachment: small file (inline) ───────────────────


class TestAddAttachmentSmall:
    def test_inline_base64_payload(self, client, small_file):
        """Small files should be sent as inline base64 attachments."""
        with patch.object(client, "_post", return_value={"Id": "att_001"}) as mock_post, \
             patch.object(client, "_save_id_map"):
            result = client.add_attachment(FAKE_MSG_ID, str(small_file))

        mock_post.assert_called_once()
        args = mock_post.call_args
        path_arg = args[0][0]
        payload = args[1]["json"]

        assert "/attachments" in path_arg
        assert payload["@odata.type"] == "#Microsoft.OutlookServices.FileAttachment"
        assert payload["Name"] == "small.txt"
        assert payload["ContentBytes"] == base64.b64encode(b"Hello, World!").decode()
        assert result == {"Id": "att_001"}

    def test_content_type_detection(self, client, tmp_path):
        """MIME type should be guessed from file extension."""
        pdf = tmp_path / "report.pdf"
        pdf.write_bytes(b"%PDF-1.4")

        with patch.object(client, "_post", return_value={}) as mock_post, \
             patch.object(client, "_save_id_map"):
            client.add_attachment(FAKE_MSG_ID, str(pdf))

        payload = mock_post.call_args[1]["json"]
        assert payload["ContentType"] == "application/pdf"

    def test_unknown_extension_fallback(self, client, tmp_path):
        """Unknown extensions should fall back to application/octet-stream."""
        f = tmp_path / "data.xyz123"
        f.write_bytes(b"\x00\x01\x02")

        with patch.object(client, "_post", return_value={}) as mock_post, \
             patch.object(client, "_save_id_map"):
            client.add_attachment(FAKE_MSG_ID, str(f))

        payload = mock_post.call_args[1]["json"]
        assert payload["ContentType"] == "application/octet-stream"


# ── add_attachment: large file (upload session) ───────────


class TestAddAttachmentLarge:
    def test_creates_upload_session(self, client, large_file):
        """Large files should trigger an upload session."""
        file_size = large_file.stat().st_size

        with patch.object(client, "_post", return_value={"uploadUrl": "https://upload.example.com/session"}) as mock_post, \
             patch("outlook_cli.client.httpx.put", return_value=MagicMock(status_code=200, content=b'{}', json=lambda: {})) as mock_put, \
             patch.object(client, "_save_id_map"):
            client.add_attachment(FAKE_MSG_ID, str(large_file))

        # Should create upload session
        session_call = mock_post.call_args
        payload = session_call[1]["json"]
        assert payload["AttachmentItem"]["attachmentType"] == "file"
        assert payload["AttachmentItem"]["name"] == "large.bin"
        assert payload["AttachmentItem"]["size"] == file_size

        # Should upload via PUT
        assert mock_put.called
        put_call = mock_put.call_args
        assert "Content-Range" in put_call[1]["headers"]

    def test_upload_session_chunking(self, client, tmp_path):
        """Files larger than chunk size should be uploaded in multiple chunks."""
        chunk_size = 4 * 1024 * 1024  # 4 MB
        file_size = chunk_size + 1024  # Just over one chunk
        f = tmp_path / "big.dat"
        f.write_bytes(b"\x00" * file_size)

        put_responses = [
            MagicMock(status_code=200, content=b"", json=lambda: {}),
            MagicMock(status_code=201, content=b'{"Id":"att"}', json=lambda: {"Id": "att"}),
        ]

        with patch.object(client, "_post", return_value={"uploadUrl": "https://upload.example.com/session"}), \
             patch("outlook_cli.client.httpx.put", side_effect=put_responses) as mock_put, \
             patch.object(client, "_save_id_map"):
            client.add_attachment(FAKE_MSG_ID, str(f))

        assert mock_put.call_count == 2
        # First chunk: 0 to chunk_size-1
        first_range = mock_put.call_args_list[0][1]["headers"]["Content-Range"]
        assert first_range == f"bytes 0-{chunk_size - 1}/{file_size}"
        # Second chunk: chunk_size to file_size-1
        second_range = mock_put.call_args_list[1][1]["headers"]["Content-Range"]
        assert second_range == f"bytes {chunk_size}-{file_size - 1}/{file_size}"


# ── add_attachment: error cases ───────────────────────────


class TestAddAttachmentErrors:
    def test_file_not_found(self, client):
        """Should raise FileNotFoundError for missing files."""
        with pytest.raises(FileNotFoundError, match="nonexistent.txt"):
            client.add_attachment("some_id", "/tmp/nonexistent.txt")

    def test_resolve_id_called(self, client, small_file):
        """Should resolve display number to real ID."""
        client._id_map["5"] = "AAMk_real_id_long_string_for_testing_purposes_here_abc123"
        with patch.object(client, "_post", return_value={}) as mock_post, \
             patch.object(client, "_save_id_map"):
            client.add_attachment("5", str(small_file))

        path_arg = mock_post.call_args[0][0]
        assert "AAMk_real_id_long_string_for_testing_purposes_here_abc123" in path_arg


# ── attach_files ──────────────────────────────────────────


class TestAttachFiles:
    def test_attaches_multiple_files(self, client, tmp_path):
        """Should call add_attachment for each file."""
        f1 = tmp_path / "a.txt"
        f2 = tmp_path / "b.txt"
        f1.write_text("aaa")
        f2.write_text("bbb")

        with patch.object(client, "add_attachment") as mock_add:
            client.attach_files("msg_id", [str(f1), str(f2)])

        assert mock_add.call_count == 2
        mock_add.assert_any_call("msg_id", str(f1))
        mock_add.assert_any_call("msg_id", str(f2))

    def test_empty_list_is_noop(self, client):
        """Empty file list should do nothing."""
        with patch.object(client, "add_attachment") as mock_add:
            client.attach_files("msg_id", [])

        mock_add.assert_not_called()


# ── create_forward_draft ──────────────────────────────────


class TestCreateForwardDraft:
    def test_creates_forward_draft(self, client):
        """Should POST to createforward, then PATCH with HTML body when comment given."""
        post_response = {
            "Id": "AAMk_forward_draft_long_id_string_for_testing_purposes_123",
            "Subject": "Fwd: Original Subject",
            "Body": {"ContentType": "HTML", "Content": "<html><body>original</body></html>"},
        }
        patch_response = {
            "Id": "AAMk_forward_draft_long_id_string_for_testing_purposes_123",
            "Subject": "Fwd: Original Subject",
        }
        client._id_map["10"] = "AAMk_original_msg_long_id_string_for_testing_purposes_here"
        with patch.object(client, "_post", return_value=post_response) as mock_post, \
             patch.object(client, "_patch", return_value=patch_response) as mock_patch, \
             patch.object(client, "_save_id_map"):
            result = client.create_forward_draft("10", ["user@example.com"], comment="FYI")

        path_arg = mock_post.call_args[0][0]
        assert "createforward" in path_arg
        payload = mock_post.call_args[1]["json"]
        assert payload["ToRecipients"][0]["EmailAddress"]["Address"] == "user@example.com"
        # Comment is now applied via PATCH as HTML, not via Comment field
        mock_patch.assert_called_once()
        patch_payload = mock_patch.call_args[1]["json"]
        assert patch_payload["Body"]["ContentType"] == "HTML"
        assert "FYI" in patch_payload["Body"]["Content"]
        assert result.subject == "Fwd: Original Subject"


# ── Threshold constant ────────────────────────────────────


class TestAttachmentThreshold:
    def test_threshold_is_3mb(self):
        assert ATTACHMENT_SIZE_THRESHOLD == 3 * 1024 * 1024

    def test_file_at_threshold_uses_upload_session(self, client, tmp_path):
        """File exactly at 3 MB should use upload session."""
        f = tmp_path / "exact.bin"
        f.write_bytes(b"\x00" * ATTACHMENT_SIZE_THRESHOLD)

        with patch.object(client, "_post", return_value={"uploadUrl": "https://upload.example.com/session"}) as mock_post, \
             patch("outlook_cli.client.httpx.put", return_value=MagicMock(status_code=200, content=b'{}', json=lambda: {})), \
             patch.object(client, "_save_id_map"):
            client.add_attachment(FAKE_MSG_ID, str(f))

        # Should use upload session (createuploadsession), not inline
        path_arg = mock_post.call_args[0][0]
        assert "createuploadsession" in path_arg

    def test_file_under_threshold_uses_inline(self, client, tmp_path):
        """File just under 3 MB should use inline base64."""
        f = tmp_path / "under.bin"
        f.write_bytes(b"\x00" * (ATTACHMENT_SIZE_THRESHOLD - 1))

        with patch.object(client, "_post", return_value={}) as mock_post, \
             patch.object(client, "_save_id_map"):
            client.add_attachment(FAKE_MSG_ID, str(f))

        path_arg = mock_post.call_args[0][0]
        assert path_arg.endswith("/attachments")
        assert "createuploadsession" not in path_arg


# ── Helper functions ──────────────────────────────────────


class TestFormatFileSize:
    def test_bytes(self):
        from outlook_cli.commands.mail import _format_file_size
        assert _format_file_size(500) == "500 B"

    def test_kilobytes(self):
        from outlook_cli.commands.mail import _format_file_size
        assert _format_file_size(2048) == "2.0 KB"

    def test_megabytes(self):
        from outlook_cli.commands.mail import _format_file_size
        assert _format_file_size(3 * 1024 * 1024) == "3.0 MB"
