"""Scope fairness and privacy checks for the offline benchmark deliverable."""
from __future__ import annotations

import json

import pytest

from outlook_cli.index_store import MailIndex
from outlook_cli.mail_store import MailStore
from scripts.benchmark_local import comparable_scope, open_legacy


def record(identity):
    return {"id": identity, "subject": "Toplantı çalışma poliçe sözleşme güvenlik görüşme yenileme",
            "body": "İş sürekliliği toplantı çalışma güvenlik görüşme yenileme.", "body_type": "Text",
            "received": "2026-10-09T10:00:00Z", "sender": {"name": "Private Name", "address": "private@example.invalid"},
            "to": [], "cc": [], "conversation_id": "private-conversation", "preview": "Toplantı çalışma"}


def add_new(store, identity, name, records):
    store.begin_sync("rest", identity, name, full=True)
    store.apply_page("rest", identity, name, records, complete=True)


@pytest.fixture
def stores(tmp_path):
    old_path = tmp_path / "legacy.sqlite3"
    old = MailIndex(old_path)
    new = MailStore(tmp_path / "new" / "mail.sqlite3")
    yield old, new, old_path
    old.close()
    new.close()


def test_comparable_scope_matches_identity_and_folded_name_alias_and_excludes_extra_folder(stores):
    old, new, old_path = stores
    old.apply("rest", "private-retained-id", "Private Original", [record("private-message-one")], full=True)
    old.apply("rest", "private-old-alias", "İŞ Çalışmaları", [record("private-message-two")], full=True)
    add_new(new, "private-retained-id", "Private Renamed", [record("private-message-one")])
    add_new(new, "private-new-alias", "is calismalari", [record("private-message-two")])
    add_new(new, "private-extra-folder", "Private Extra", [record("private-extra-message")])
    readonly = open_legacy(old_path)
    try:
        result = comparable_scope(readonly, new, backend="rest", repeats=2)
    finally:
        readonly.close()
    assert result["ok"]
    assert result["scope_comparison"] == "matched"
    assert result["legacy_folder_count"] == result["mapped_folder_count"] == 2
    assert result["legacy_message_count"] == result["new_message_count"] == 2
    assert new.query("toplantı", match_mode="exact")[1]["total_matches"] == 3
    assert result["queries"][0]["legacy"]["total_matches_sum"] == 2
    assert result["queries"][0]["new"]["total_matches_sum"] == 2
    assert result["new"]["folder_query_measurements"] == 32
    encoded = json.dumps(result, ensure_ascii=False)
    for secret in ("private-retained-id", "private-old-alias", "private-new-alias", "Private Original",
                   "İŞ Çalışmaları", "private-message-one", "private@example.invalid", "Private Name"):
        assert secret not in encoded


def test_missing_scope_mapping_is_unknown_and_does_not_measure_partial_subset(stores):
    old, new, _ = stores
    old.apply("rest", "private-found", "Private Found", [record("private-one")], full=True)
    old.apply("rest", "private-missing", "Private Missing", [record("private-two")], full=True)
    add_new(new, "private-found", "Private Found", [record("private-one")])
    result = comparable_scope(old, new, backend="rest", repeats=1)
    assert not result["ok"]
    assert result["scope_comparison"] == "unknown"
    assert result["unmapped_folder_count"] == 1
    assert result["new_message_count"] is None
    assert result["error"]["code"] == "scope_mapping_unknown"
    assert "queries" not in result
    assert "private-missing" not in json.dumps(result)


def test_ambiguous_folded_folder_name_is_unknown_even_if_message_counts_match(stores):
    old, new, _ = stores
    old.apply("rest", "private-old", "İŞ", [record("private-one")], full=True)
    add_new(new, "private-new-a", "iş", [record("private-one")])
    add_new(new, "private-new-b", "IS", [])
    result = comparable_scope(old, new, backend="rest", repeats=1)
    assert result["scope_comparison"] == "unknown"
    assert result["mapped_folder_count"] == 0
    assert not result["ok"]


def test_scope_alias_cannot_reuse_authoritative_identity_match(stores):
    old, new, _ = stores
    old.apply("rest", "private-id", "Original", [record("private-one")], full=True)
    old.apply("rest", "private-alias", "İŞ", [record("private-two")], full=True)
    add_new(new, "private-id", "is", [record("private-one")])
    result = comparable_scope(old, new, backend="rest", repeats=1)
    assert result["scope_comparison"] == "unknown"
    assert result["mapped_folder_count"] == 1
    assert not result["ok"]
