"""CLI integration tests for remaining command modules."""

from __future__ import annotations

import base64
import json

from outlook_cli import category_manager, signature_manager
from outlook_cli.commands import attachments, auth as auth_cmd, categories, contacts, folders, manage, open_item, search, signatures


def test_login_command_reports_success(runner, tty_mode, monkeypatch):
    messages = []
    monkeypatch.setattr(auth_cmd, "do_login", lambda force=False, debug=False, **kwargs: "token")
    monkeypatch.setattr(auth_cmd, "verify_token", lambda token: True)
    monkeypatch.setattr(auth_cmd, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(auth_cmd.login, ["--force"])

    assert result.exit_code == 0
    assert messages == ["Logged in successfully for account 'default'. Token cached."]


def test_login_command_exits_on_runtime_error(runner, tty_mode, monkeypatch):
    errors = []
    monkeypatch.setattr(auth_cmd, "do_login", lambda force=False, debug=False, **kwargs: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(auth_cmd, "print_error", lambda msg: errors.append(msg))

    result = runner.invoke(auth_cmd.login, [])

    assert result.exit_code == 1
    assert errors == ["boom"]


def test_whoami_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_me": lambda self: {"DisplayName": "Alice"}})()
    monkeypatch.setattr(auth_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setattr(auth_cmd, "get_account_name", lambda account_name=None: "default")

    result = runner.invoke(auth_cmd.whoami, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["DisplayName"] == "Alice"
    assert payload["data"]["AccountProfile"] == "default"


def test_folders_command_can_export_json(runner, tty_mode, monkeypatch, make_folder, tmp_path):
    fake_client = type("Client", (), {"get_folders": lambda self: [make_folder(name="Inbox")]})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)

    output = tmp_path / "folders.json"
    result = runner.invoke(folders.folders, ["--json", "--output", str(output)])

    assert result.exit_code == 0
    assert json.loads(output.read_text())[0]["name"] == "Inbox"


def test_folder_command_outputs_json(runner, tty_mode, monkeypatch, make_email):
    fake_client = type("Client", (), {"get_messages": lambda self, **kwargs: [make_email(subject="In folder")]})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)

    result = runner.invoke(folders.folder, ["Inbox", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "In folder"


def test_search_command_reports_no_results(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"search_messages": lambda self, query, top=25: []})()
    messages = []
    monkeypatch.setattr(search, "_get_client", lambda: fake_client)
    monkeypatch.setattr(search, "print_error", lambda msg: messages.append(msg))

    result = runner.invoke(search.search, ["invoice"])

    assert result.exit_code == 0
    assert messages == ["No results found."]


def test_contacts_command_outputs_json(runner, tty_mode, monkeypatch, make_contact):
    fake_client = type("Client", (), {"get_contacts": lambda self, top=50: [make_contact()]})()
    monkeypatch.setattr(contacts, "_get_client", lambda: fake_client)

    result = runner.invoke(contacts.contacts, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["display_name"] == "Alice Smith"


def test_categories_command_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = type(
        "Client",
        (),
        {"get_master_categories": lambda self: {"Body": {"CategoryDetailsList": [{"Category": "Finance"}]}}},
    )()
    monkeypatch.setattr(categories, "_get_client", lambda: fake_client)

    result = runner.invoke(categories.categories, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["Category"] == "Finance"


def test_categorize_and_uncategorize_loop_over_ids(runner, tty_mode, monkeypatch):
    class FakeClient:
        def add_category(self, message_id, category):
            self.added = getattr(self, "added", []) + [(message_id, category)]
            return ["Finance"]

        def remove_category(self, message_id, category):
            self.removed = getattr(self, "removed", []) + [(message_id, category)]
            return []

    fake_client = FakeClient()
    monkeypatch.setattr(categories, "_get_client", lambda: fake_client)

    result_add = runner.invoke(categories.categorize, ["1", "2", "Finance"])
    result_remove = runner.invoke(categories.uncategorize, ["1", "2", "Finance"])

    assert result_add.exit_code == 0
    assert result_remove.exit_code == 0
    assert fake_client.added == [("1", "Finance"), ("2", "Finance")]
    assert fake_client.removed == [("1", "Finance"), ("2", "Finance")]


def test_category_management_commands_delegate_to_manager(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(categories, "get_token", lambda: "token")
    monkeypatch.setattr(category_manager, "rename_category", lambda *args, **kwargs: 3)
    monkeypatch.setattr(category_manager, "clear_category", lambda *args, **kwargs: 4)
    monkeypatch.setattr(category_manager, "delete_category", lambda *args, **kwargs: None)
    monkeypatch.setattr(category_manager, "create_category", lambda *args, **kwargs: None)

    rename = runner.invoke(categories.category_rename, ["Old", "New", "--no-propagate"])
    clear = runner.invoke(categories.category_clear, ["Finance", "--folder", "Inbox", "--max", "2", "-y"])
    delete = runner.invoke(categories.category_delete, ["Finance", "--no-propagate", "-y"])
    create = runner.invoke(categories.category_create, ["NewCat", "--color", "7"])

    assert rename.exit_code == 0
    assert clear.exit_code == 0
    assert delete.exit_code == 0
    assert create.exit_code == 0


def test_attachments_command_downloads_inline_and_remote_content(runner, tty_mode, monkeypatch, tmp_path, make_attachment):
    inline = make_attachment(name="inline.txt", content_bytes=base64.b64encode(b"inline").decode())
    remote = make_attachment(id="att-2", name="remote.txt", content_bytes=None)

    class FakeClient:
        def get_attachments(self, message_id):
            return [inline, remote]

        def download_attachment(self, message_id, attachment_id):
            return make_attachment(id=attachment_id, name="remote.txt", content_bytes=base64.b64encode(b"remote").decode())

    monkeypatch.setattr(attachments, "_get_client", lambda: FakeClient())
    monkeypatch.setattr(attachments, "print_attachments", lambda atts: None)

    result = runner.invoke(attachments.attachments, ["1", "--download", "--save-to", str(tmp_path)])

    assert result.exit_code == 0
    assert (tmp_path / "inline.txt").read_bytes() == b"inline"
    assert (tmp_path / "remote.txt").read_bytes() == b"remote"


def test_signature_commands_delegate_to_manager(runner, tty_mode, monkeypatch, tmp_path):
    monkeypatch.setattr(signatures, "get_token", lambda: "token")
    monkeypatch.setattr(signature_manager, "pull_signature", lambda token: ("<b>sig</b>", "Sent mail"))
    monkeypatch.setattr(signature_manager, "save_signature", lambda name, html: tmp_path / f"{name}.html")
    monkeypatch.setattr(signature_manager, "list_signatures", lambda: ["default"])
    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(signature_manager, "delete_signature", lambda name: None)
    monkeypatch.setitem(signatures.cfg, "default_signature", "default")
    monkeypatch.setattr(signatures.click, "prompt", lambda *args, **kwargs: "default")

    pull = runner.invoke(signatures.signature_pull, [])
    list_result = runner.invoke(signatures.signature_list, [])
    show = runner.invoke(signatures.signature_show, ["default"])
    delete = runner.invoke(signatures.signature_delete, ["default", "-y"])

    assert pull.exit_code == 0
    assert list_result.exit_code == 0
    assert show.exit_code == 0
    assert delete.exit_code == 0


def test_manage_commands_delegate_to_client(runner, tty_mode, monkeypatch):
    class FakeClient:
        def mark_read(self, message_id, is_read=True):
            self.marked = getattr(self, "marked", []) + [(message_id, is_read)]

        def move_message(self, message_id, destination):
            self.moved = getattr(self, "moved", []) + [(message_id, destination)]

        def copy_message(self, message_id, destination):
            self.copied = getattr(self, "copied", []) + [(message_id, destination)]

        def delete_message(self, message_id):
            self.deleted = getattr(self, "deleted", []) + [message_id]

        def set_flag(self, message_id, status="flagged", due_date=None):
            self.flagged = getattr(self, "flagged", []) + [(message_id, status, due_date)]

        def pin_message(self, message_id, pinned=True):
            self.pinned = getattr(self, "pinned", []) + [(message_id, pinned)]

    fake_client = FakeClient()
    monkeypatch.setattr(manage, "_get_client", lambda: fake_client)

    mark = runner.invoke(manage.mark_read, ["1", "2", "--unread"])
    move_result = runner.invoke(manage.move, ["1", "2", "Archive"])
    copy_result = runner.invoke(manage.copy, ["1", "2", "Finance"])
    delete_result = runner.invoke(manage.delete, ["1", "2"], input="y\n")
    flag_result = runner.invoke(manage.flag, ["1", "2", "--due", "2026-03-20"])
    pin_result = runner.invoke(manage.pin, ["1", "2", "--unpin"])

    assert mark.exit_code == 0
    assert move_result.exit_code == 0
    assert copy_result.exit_code == 0
    assert delete_result.exit_code == 0
    assert flag_result.exit_code == 0
    assert pin_result.exit_code == 0
    assert fake_client.marked == [("1", False), ("2", False)]
    assert fake_client.moved == [("1", "Archive"), ("2", "Archive")]
    assert fake_client.copied == [("1", "Finance"), ("2", "Finance")]
    assert fake_client.deleted == ["1", "2"]
    assert fake_client.flagged == [("1", "flagged", "2026-03-20"), ("2", "flagged", "2026-03-20")]
    assert fake_client.pinned == [("1", False), ("2", False)]


def test_open_command_opens_browser(runner, tty_mode, monkeypatch):
    class FakeClient:
        def get_open_target(self, item_id):
            self.called = item_id
            return ("message", "https://outlook.office365.com/owa/?ItemID=abc")

    fake_client = FakeClient()
    opened = []
    messages = []
    monkeypatch.setattr(open_item, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(open_item.webbrowser, "open", lambda url: opened.append(url) or True)
    monkeypatch.setattr(open_item, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(open_item.open_item, ["3"])

    assert result.exit_code == 0
    assert fake_client.called == "3"
    assert opened == ["https://outlook.office365.com/owa/?ItemID=abc"]
    assert messages == ["Opened message #3 in browser"]


def test_open_command_can_print_url_without_opening_browser(runner, tty_mode, monkeypatch):
    class FakeClient:
        def get_open_target(self, item_id):
            self.called = item_id
            return ("event", "https://outlook.office365.com/owa/?itemid=evt")

    fake_client = FakeClient()
    monkeypatch.setattr(open_item, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(open_item.webbrowser, "open", lambda url: (_ for _ in ()).throw(AssertionError("should not open browser")))

    result = runner.invoke(open_item.open_item, ["42", "--print-url"])

    assert result.exit_code == 0
    assert fake_client.called == "42"
    assert result.output.strip() == "https://outlook.office365.com/owa/?itemid=evt"
