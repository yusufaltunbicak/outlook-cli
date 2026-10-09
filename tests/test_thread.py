"""Tests for thread feature — get_thread client method and print_thread formatter."""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from outlook_cli.formatter import print_thread
from outlook_cli.models import Email, EmailAddress


def _make_email(subject: str, sender: str, num: int, is_read: bool = True) -> Email:
    """Helper to create a minimal Email for testing."""
    from datetime import datetime, timezone
    return Email(
        id=f"id_{num}",
        subject=subject,
        sender=EmailAddress(name=sender, address=f"{sender.lower().replace(' ', '.')}@contoso.com"),
        to=[EmailAddress(name="Bob", address="bob@contoso.com")],
        cc=[],
        received=datetime(2026, 3, 10, 9, 0, tzinfo=timezone.utc),
        preview=f"Preview of {subject}",
        body=f"Body of {subject}",
        body_type="Text",
        is_read=is_read,
        has_attachments=False,
        importance="Normal",
        conversation_id="conv_abc123",
        display_num=num,
    )


class TestPrintThread:
    def test_shows_all_messages(self, capsys):
        messages = [
            _make_email("Initial question", "Alice", 1),
            _make_email("Re: Initial question", "Bob", 2),
            _make_email("Re: Re: Initial question", "Alice", 3),
        ]
        print_thread(messages)
        captured = capsys.readouterr()
        assert "3 messages" in captured.err  # Rich goes to stderr
        assert "Alice" in captured.err
        assert "Bob" in captured.err

    def test_unread_marker(self, capsys):
        messages = [
            _make_email("Hello", "Alice", 1, is_read=True),
            _make_email("Re: Hello", "Bob", 2, is_read=False),
        ]
        print_thread(messages)
        captured = capsys.readouterr()
        assert "*" in captured.err  # unread marker for Bob's message

    def test_single_message_thread(self):
        """A single-message thread should still work."""
        messages = [_make_email("Solo message", "Alice", 1)]
        # Should not raise
        print_thread(messages)

    def test_long_body_truncation(self, capsys):
        """Messages with >20 lines should be truncated."""
        email = _make_email("Long email", "Alice", 1)
        email.body = "\n".join(f"Line {i}" for i in range(50))
        print_thread([email])
        captured = capsys.readouterr()
        assert "more lines" in captured.err


class TestGetThread:
    def test_returns_conversation_messages(self):
        """get_thread should search by subject and filter by ConversationId client-side."""
        from outlook_cli.client import OutlookClient

        with patch.object(OutlookClient, '__init__', lambda self, token: None):
            client = OutlookClient.__new__(OutlookClient)
            client._id_map = {"1": "real_id_1"}
            client._next_num = 10

            mock_email = _make_email("Re: Test", "Alice", 1)
            mock_email.id = "msg_2"
            client.get_message = MagicMock(return_value=mock_email)

            # API returns messages from search — includes same + different conversations
            api_response = {
                "value": [
                    {
                        "Id": "msg_1",
                        "Subject": "Test",
                        "ConversationId": "conv_abc123",
                        "ReceivedDateTime": "2026-03-10T09:00:00Z",
                    },
                    {
                        "Id": "msg_2",
                        "Subject": "Re: Test",
                        "ConversationId": "conv_abc123",
                        "ReceivedDateTime": "2026-03-10T10:00:00Z",
                    },
                    {
                        "Id": "msg_other",
                        "Subject": "Test something else",
                        "ConversationId": "conv_DIFFERENT",
                        "ReceivedDateTime": "2026-03-10T11:00:00Z",
                    },
                ]
            }
            client._get = MagicMock(return_value=api_response)
            client._assign_display_nums = MagicMock()

            messages = client.get_thread("1")

            # Should only include messages with matching ConversationId
            assert len(messages) == 2
            assert messages[0].subject == "Test"
            assert messages[1].subject == "Re: Test"

            # Verify $search was used (not $filter)
            call_args = client._get.call_args
            assert "$search" in call_args[1]["params"]

    def test_no_conversation_id_returns_single(self):
        """If message has no conversation_id, return just that message."""
        from outlook_cli.client import OutlookClient

        with patch.object(OutlookClient, '__init__', lambda self, token: None):
            client = OutlookClient.__new__(OutlookClient)
            client._id_map = {"1": "real_id_1"}
            client._next_num = 10

            mock_email = _make_email("Solo", "Alice", 1)
            mock_email.conversation_id = ""
            client.get_message = MagicMock(return_value=mock_email)

            messages = client.get_thread("1")
            assert len(messages) == 1
            assert messages[0].subject == "Solo"
