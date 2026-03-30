"""CLI integration tests for mail and scheduling commands."""

from __future__ import annotations

import json
from unittest.mock import MagicMock

from outlook_cli import cli as cli_module
from outlook_cli.commands import mail, schedule
from outlook_cli import signature_manager


def test_send_with_attachments_uses_draft_flow(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "report.txt"
    attachment.write_text("hello")
    fake_client = MagicMock()
    fake_client.create_draft.return_value = make_email(id="draft-1")
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)
    monkeypatch.setitem(mail.cfg, "default_signature", "default")
    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(signature_manager, "append_signature", lambda body, sig, is_html: ("Body+Sig", True))

    result = runner.invoke(
        mail.send,
        ["a@example.com", "Subject", "Body", "--attach", str(attachment), "-y"],
    )

    assert result.exit_code == 0
    fake_client.create_draft.assert_called_once_with(
        to=["a@example.com"],
        subject="Subject",
        body="Body+Sig",
        cc=None,
        html=True,
    )
    fake_client.attach_files.assert_called_once_with("draft-1", [str(attachment)])
    fake_client.send_draft.assert_called_once_with("draft-1")


def test_send_without_attachments_calls_send_mail(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.send, ["a@example.com,b@example.com", "Subject", "Body", "-y"])

    assert result.exit_code == 0
    fake_client.send_mail.assert_called_once_with(
        to=["a@example.com", "b@example.com"],
        subject="Subject",
        body="Body",
        cc=None,
        html=False,
    )


def test_send_reads_body_from_file(runner, tty_mode, monkeypatch, tmp_path):
    body_file = tmp_path / "body.txt"
    body_file.write_text("Body from file")
    fake_client = MagicMock()
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.send, ["a@example.com", "Subject", "--body-file", str(body_file), "-y"])

    assert result.exit_code == 0
    fake_client.send_mail.assert_called_once_with(
        to=["a@example.com"],
        subject="Subject",
        body="Body from file",
        cc=None,
        html=False,
    )


def test_draft_creates_message_and_attaches_files(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "draft.txt"
    attachment.write_text("x")
    fake_client = MagicMock()
    fake_client.create_draft.return_value = make_email(id="draft-2")
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.draft, ["a@example.com", "Subject", "Body", "--attach", str(attachment)])

    assert result.exit_code == 0
    fake_client.attach_files.assert_called_once_with("draft-2", [str(attachment)])


def test_draft_send_confirms_and_sends(runner, tty_mode, monkeypatch, make_email):
    fake_client = MagicMock()
    fake_client.get_message.return_value = make_email(display_num=4)
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)

    result = runner.invoke(mail.draft_send, ["4"], input="y\n")

    assert result.exit_code == 0
    fake_client.send_draft.assert_called_once_with("4")


def test_read_marks_unread_message_as_read(runner, tty_mode, monkeypatch, make_email):
    fake_client = MagicMock()
    fake_client.get_message.return_value = make_email(is_read=False)
    shown = []
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)
    monkeypatch.setattr(mail, "print_email", lambda email: shown.append(email.id))

    result = runner.invoke(mail.read, ["1"])

    assert result.exit_code == 0
    assert shown == ["msg-1"]
    fake_client.mark_read.assert_called_once_with("1")


def test_reply_with_attachments_uses_reply_draft_flow(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "reply.txt"
    attachment.write_text("x")
    fake_client = MagicMock()
    fake_client.create_reply_draft.return_value = make_email(id="reply-draft")
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)

    result = runner.invoke(mail.reply, ["3", "Thanks", "--attach", str(attachment), "-y"])

    assert result.exit_code == 0
    fake_client.create_reply_draft.assert_called_once_with("3", comment="Thanks", reply_all=False)
    fake_client.attach_files.assert_called_once_with("reply-draft", [str(attachment)])
    fake_client.send_draft.assert_called_once_with("reply-draft")


def test_reply_reads_body_from_stdin(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)

    result = runner.invoke(mail.reply, ["3", "--body-file", "-", "-y"], input="Thanks from stdin")

    assert result.exit_code == 0
    fake_client.reply.assert_called_once_with("3", "Thanks from stdin", reply_all=False)


def test_reply_draft_outputs_json(runner, tty_mode, monkeypatch, make_email):
    fake_client = MagicMock()
    fake_client.create_reply_draft.return_value = make_email(id="reply-draft", subject="Re: Topic")
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.reply_draft, ["3", "Thanks", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["data"]["subject"] == "Re: Topic"


def test_forward_with_attachments_uses_forward_draft_flow(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "forward.txt"
    attachment.write_text("x")
    fake_client = MagicMock()
    fake_client.create_forward_draft.return_value = make_email(id="forward-draft")
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)

    result = runner.invoke(mail.forward, ["8", "a@example.com", "--attach", str(attachment), "-y"])

    assert result.exit_code == 0
    fake_client.create_forward_draft.assert_called_once_with("8", ["a@example.com"], comment="")
    fake_client.attach_files.assert_called_once_with("forward-draft", [str(attachment)])
    fake_client.send_draft.assert_called_once_with("forward-draft")


def test_parse_schedule_time_accepts_relative_offsets():
    dt = schedule._parse_schedule_time("+1h30m")
    now = schedule.datetime.now(schedule.timezone.utc)

    assert 89 * 60 <= (dt - now).total_seconds() <= 91 * 60


def test_schedule_with_attachments_uses_draft_schedule_flow(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "schedule.txt"
    attachment.write_text("x")
    fake_client = MagicMock()
    fake_client.create_draft.return_value = make_email(id="draft-3")
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Planned", "Body", "+30m", "--attach", str(attachment), "-y"],
    )

    assert result.exit_code == 0
    fake_client.create_draft.assert_called_once()
    fake_client.attach_files.assert_called_once_with("draft-3", [str(attachment)])
    assert fake_client.schedule_draft.call_count == 1


def test_send_rejects_body_and_body_file_together(runner, tty_mode, monkeypatch, tmp_path):
    body_file = tmp_path / "body.txt"
    body_file.write_text("Body from file")
    fake_client = MagicMock()
    monkeypatch.setattr(mail, "_get_client", lambda: fake_client)

    result = runner.invoke(mail.send, ["a@example.com", "Subject", "Body", "--body-file", str(body_file), "-y"])

    assert result.exit_code == 2
    assert "Use either BODY or --body-file, not both." in result.output
    fake_client.send_mail.assert_not_called()


def test_schedule_reads_body_from_stdin_in_dry_run_json(runner, monkeypatch):
    result = runner.invoke(
        cli_module.cli,
        [
            "schedule",
            "a@example.com",
            "Subject",
            "+30m",
            "--body-file",
            "-",
            "--dry-run",
            "--json",
        ],
        input="Scheduled from stdin",
    )

    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["data"]["dry_run"] is True
    assert payload["data"]["request"]["body"] == "Scheduled from stdin"


def test_schedule_list_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = [{"subject": "Planned", "scheduled_at": "2026-03-20T10:00:00Z"}]
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule_list, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Planned"


def test_schedule_cancel_reports_invalid_index(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = []
    errors = []
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setattr(schedule, "print_error", lambda msg: errors.append(msg))

    result = runner.invoke(schedule.schedule_cancel, ["3", "-y"])

    assert result.exit_code == 0
    assert errors == ["Invalid index #3. Run 'outlook schedule-list' to see entries."]
    fake_client.cancel_scheduled_entry.assert_not_called()


def test_schedule_draft_confirms_and_calls_client(runner, tty_mode, monkeypatch, make_email):
    fake_client = MagicMock()
    fake_client.get_message.return_value = make_email(display_num=5)
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule_draft, ["5", "+1h"], input="y\n")

    assert result.exit_code == 0
    assert fake_client.schedule_draft.call_count == 1
