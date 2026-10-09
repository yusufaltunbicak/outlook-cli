"""Offline command/output/safety contracts for local mail research."""
from __future__ import annotations

import json
from contextlib import closing
from pathlib import Path
from types import SimpleNamespace

import pytest

from outlook_cli import account
from outlook_cli.cli import cli
from outlook_cli.commands import local as local_cmd
from outlook_cli.exceptions import RateLimitError
from outlook_cli.mail_store import MailStore


@pytest.fixture
def mailbox_store():
    path = local_cmd.store_path("default")
    legacy = path.parent / "index.sqlite3"
    records = [
        {"id": "message-a", "subject": "Şirket toplantısı", "sender": {"name": "Şule Işık", "address": "alice@example.com"},
         "to": [{"name": "Bob", "address": "bob@example.net"}], "cc": [], "received": "2026-10-05T09:00:00Z",
         "body": "<html><body><p>Şirket toplantısında ölçüm görüşüldü.</p><script>unwanted-script()</script></body></html>",
         "body_type": "HTML", "preview": "Şirket toplantısı", "conversation_id": "thread-a", "is_read": False,
         "has_attachments": True, "attachments": [{"id": "attachment-a", "name": "rapor.pdf", "size": 123,
                                                    "content_type": "application/pdf", "is_inline": False}]},
        {"id": "message-c", "subject": "Başka çalışma", "sender": {"name": "Şule Işık", "address": "alice@example.com"},
         "to": [{"address": "other@example.net"}], "cc": [], "received": "2026-10-04T09:00:00Z",
         "body": "Bağımsız konu", "body_type": "Text", "preview": "Bağımsız konu", "conversation_id": "thread-c",
         "is_read": True, "has_attachments": False},
    ]
    with closing(MailStore(path)) as store:
        store.begin_sync("rest", "inbox", "Gelen Kutusu", full=True, options_signature="fixture")
        store.apply_page("rest", "inbox", "Gelen Kutusu", records, complete=True, cursor="private-cursor-inbox")
        store.begin_sync("rest", "sent", "Gönderilen Öğeler", full=True, options_signature="fixture")
        store.apply_page("rest", "sent", "Gönderilen Öğeler", [
            {"id": "message-b", "subject": "Ynt: Şirket toplantısı", "sender": {"name": "Bob", "address": "bob@example.net"},
             "to": [{"address": "alice@example.com"}], "cc": [], "received": "2026-10-06T09:00:00Z",
             "body": "Toplantı yanıtı", "body_type": "Text", "preview": "Toplantı yanıtı", "conversation_id": "thread-a",
             "is_read": True, "has_attachments": False}], complete=True, cursor="private-cursor-sent")
        store.set_inventory("rest", {"inbox", "sent"})
    legacy.write_bytes(b"legacy database retained")
    return path, legacy


def parsed(result):
    assert result.exit_code == 0, result.stdout
    payload = json.loads(result.stdout)
    assert payload["ok"] is True
    assert payload["schema_version"] == "1"
    return payload


@pytest.mark.parametrize("arguments", [
    ["status"], ["search", "toplanti"], ["read", "message-a"], ["thread", "message-a"],
    ["related", "message-a"], ["attachments", "message-a"],
])
def test_local_research_commands_need_no_auth_network_or_browser(runner, mailbox_store, arguments):
    # conftest denies all real network/keychain/browser boundaries.
    result = runner.invoke(cli, ["--no-input", "local", *arguments, "--json"])
    payload = parsed(result)
    assert "private-cursor" not in result.stdout
    if arguments[0] == "status":
        assert payload["data"]["whole_mailbox_complete"] is True
        assert payload["data"]["complete"] is True
    if arguments[0] == "search":
        assert payload["meta"]["local"] is True
        assert {row["id"] for row in payload["data"]} == {"message-a", "message-b"}
        assert all("body" not in row for row in payload["data"])
        assert all(len(row["snippet"]) < 260 for row in payload["data"])
        assert any("[[" in row["highlighted"] for row in payload["data"])


def test_compound_turkish_query_and_selected_output_fields(runner, mailbox_store):
    query = 'toplanti from:alice@example.com to:bob@example.net folder:"Gelen Kutusu" after:2026-10-01 before:2026-10-07 thread:thread-a has:attachment'
    result = runner.invoke(cli, ["local", "search", query, "--fields", "id,snippet,highlighted", "--json"])
    payload = parsed(result)
    assert len(payload["data"]) == 1
    assert payload["data"][0]["id"] == "message-a"
    assert set(payload["data"][0]) == {"id", "snippet", "highlighted"}
    assert payload["meta"]["total_matches"] == 1


def test_json_file_export_matches_stdout_and_data_only_remains_explicit(runner, mailbox_store, tmp_path):
    target = tmp_path / "results.json"
    result = runner.invoke(cli, ["local", "search", "toplanti", "--json", "--output", str(target)])
    payload = parsed(result)
    assert json.loads(target.read_text()) == payload
    raw_target = tmp_path / "raw.json"
    result = runner.invoke(cli, ["local", "search", "toplanti", "--data-only", "--output", str(raw_target)])
    assert result.exit_code == 0
    assert isinstance(json.loads(result.stdout), list)
    assert json.loads(raw_target.read_text()) == json.loads(result.stdout)


def test_local_read_defaults_to_plain_text_and_never_changes_read_state(runner, mailbox_store):
    first = parsed(runner.invoke(cli, ["local", "read", "message-a", "--json"]))["data"]
    second = parsed(runner.invoke(cli, ["local", "read", "message-a", "--json"]))["data"]
    assert first["body_type"] == "Text"
    assert "<html>" not in first["body"]
    assert "unwanted-script" not in first["body"]
    assert first["is_read"] is False
    assert second["is_read"] is False
    html = parsed(runner.invoke(cli, ["local", "read", "message-a", "--html", "--json"]))["data"]
    assert "<html>" in html["body"]


@pytest.mark.parametrize("identity", ["message-a", "thread-a"])
def test_thread_crosses_folders_and_reports_limit(runner, mailbox_store, identity):
    payload = parsed(runner.invoke(cli, ["local", "thread", identity, "--limit", "1", "--json"]))
    assert payload["meta"]["has_more"] is True
    assert payload["meta"]["result_complete"] is False
    payload = parsed(runner.invoke(cli, ["local", "thread", identity, "--json"]))
    assert [row["id"] for row in payload["data"]] == ["message-a", "message-b"]
    assert payload["meta"]["result_complete"] is True


def test_related_and_attachment_research_are_bounded(runner, mailbox_store):
    related = parsed(runner.invoke(cli, ["local", "related", "message-a", "--json"]))
    assert {row["id"] for row in related["data"]} == {"message-b", "message-c"}
    assert all("body" not in row for row in related["data"])
    attachments = parsed(runner.invoke(cli, ["local", "attachments", "message-a", "--json"]))["data"]
    assert attachments == [{"id": "attachment-a", "name": "rapor.pdf", "content_type": "application/pdf", "size": 123, "is_inline": 0}]


def test_local_missing_message_retains_not_found_exit_contract(runner, mailbox_store):
    result = runner.invoke(cli, ["local", "read", "unknown", "--json", "--no-input"])
    assert result.exit_code == 5
    payload = json.loads(result.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "not_found"


def test_incomplete_folder_rejects_require_complete_without_false_whole_coverage(runner, mailbox_store):
    path, _ = mailbox_store
    with closing(MailStore(path)) as store:
        store.failed("rest", "sent", "Gönderilen Öğeler", "fixture-sync-failure")
    result = runner.invoke(cli, ["local", "search", "toplanti", "--require-complete", "--json"])
    assert result.exit_code != 0
    assert json.loads(result.stdout)["ok"] is False
    payload = parsed(runner.invoke(cli, ["local", "search", "toplanti", "--json"]))
    assert payload["meta"]["complete"] is False
    assert payload["meta"]["whole_mailbox_complete"] is False
    assert payload["meta"]["result_complete"] is False
    inbox = parsed(runner.invoke(cli, ["local", "search", "toplanti", "--folder", "inbox", "--require-complete", "--json"]))
    assert inbox["meta"]["complete"] is True
    assert inbox["meta"]["whole_mailbox_complete"] is False


def test_discovered_unsynced_folder_is_visible_in_coverage(runner, mailbox_store):
    path, _ = mailbox_store
    with closing(MailStore(path)) as store:
        store.set_inventory("rest", {"inbox", "sent", "unsynced-archive"})
    payload = parsed(runner.invoke(cli, ["local", "status", "--json"]))["data"]
    assert payload["whole_mailbox_complete"] is False
    assert payload["scope"] == "explicitly_synced_folders"
    assert payload["discovered_folder_count"] == 3


@pytest.mark.parametrize("operation", ["sync", "purge", "compact"])
def test_dry_run_local_mutations_do_not_change_store_or_use_network(runner, mailbox_store, operation):
    path, legacy = mailbox_store
    before = path.read_bytes(), legacy.read_bytes()
    result = runner.invoke(cli, ["local", operation, "--dry-run", "--no-input", "--json"])
    payload = parsed(result)
    assert payload["data"]["dry_run"] is True
    assert (path.read_bytes(), legacy.read_bytes()) == before


def test_no_input_purge_requires_yes_and_preserves_both_stores(runner, mailbox_store):
    path, legacy = mailbox_store
    result = runner.invoke(cli, ["local", "purge", "--no-input", "--json"])
    assert result.exit_code == 2
    assert json.loads(result.stdout)["ok"] is False
    assert path.exists() and legacy.exists()


def test_purge_removes_only_new_store_and_sidecars(runner, mailbox_store):
    path, legacy = mailbox_store
    for suffix in ("-wal", "-shm"):
        Path(str(path) + suffix).write_bytes(b"local sidecar")
    payload = parsed(runner.invoke(cli, ["local", "purge", "--yes", "--no-input", "--json"]))
    assert payload["data"]["legacy_index_retained"] is True
    assert not path.exists()
    assert all(not Path(str(path) + suffix).exists() for suffix in ("-wal", "-shm"))
    assert legacy.read_bytes() == b"legacy database retained"


def test_local_schema_exposes_leaf_output_and_safety_options_offline(runner):
    payload = parsed(runner.invoke(cli, ["schema", "local", "--no-input", "--json"]))
    leaves = payload["data"]["schema"]["commands"]
    assert {"sync", "search", "read", "thread", "related", "attachments", "status", "purge", "import-index"} <= set(leaves)
    for leaf in leaves.values():
        options = {option for parameter in leaf["parameters"] for option in parameter.get("options", [])}
        assert {"--json", "--output", "--data-only"} <= options
    assert not account.get_account_paths("default").token_file.exists()


def test_sync_throttle_stops_other_folders_and_next_run_obeys_saved_cooldown(runner, mailbox_store, monkeypatch):
    requested = []
    reader = SimpleNamespace(backend="rest", requests=0, response_bytes=0)
    monkeypatch.setattr(local_cmd, "_get_client", lambda *_: object())
    monkeypatch.setattr(local_cmd, "RestReader", lambda *args, **kwargs: reader)
    monkeypatch.setattr(local_cmd, "discover_folders", lambda _: [
        {"id": "inbox", "displayName": "Gelen Kutusu"}, {"id": "sent", "displayName": "Gönderilen Öğeler"}])

    def throttled(store, reader, folder, **kwargs):
        requested.append(folder["id"])
        raise RateLimitError("fixture server throttle", retry_after=120)

    monkeypatch.setattr(local_cmd, "sync_folder", throttled)
    result = runner.invoke(cli, ["local", "sync", "--no-input", "--json"])
    assert result.exit_code == 1
    payload = json.loads(result.stdout)
    assert payload["ok"] is False
    assert requested == ["inbox"]
    assert payload["data"][0]["error"]["code"] == "rate_limited"
    assert payload["meta"]["partial"] is True
    assert payload["meta"]["remaining_folders"] == 2
    assert payload["meta"]["whole_mailbox_complete"] is False
    assert payload["meta"]["cooldown_seconds"] > 110
    assert "fixture server throttle" not in result.stdout

    def no_client(*_args):
        raise AssertionError("Cooldown must stop before authentication or API access")

    monkeypatch.setattr(local_cmd, "_get_client", no_client)
    retry = runner.invoke(cli, ["local", "sync", "--no-input", "--json"])
    assert retry.exit_code == 7
    assert json.loads(retry.stdout)["error"]["code"] == "rate_limited"
    assert requested == ["inbox"]
