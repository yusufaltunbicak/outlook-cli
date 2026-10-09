from __future__ import annotations

import hashlib
import json
import sqlite3
import unicodedata
from pathlib import Path

import pytest

from outlook_cli.exceptions import OutlookCliError, ResourceNotFoundError
from outlook_cli.index_store import MailIndex
from outlook_cli.mail_store import MailStore, fold


def message(identity="one", **changes):
    record = {"id": identity, "subject": "İŞ Görüşmesi", "sender": {"name": "Şule Öztürk", "address": "sule@example.com"},
              "to": [{"name": "Gökçe", "address": "gokce@work.test"}], "cc": [], "bcc": [],
              "received": "2026-10-09T12:00:00+03:00", "preview": "Toplantı özeti", "body": "Şirket toplantıları ve ödeme görüşmeleri.",
              "body_type": "Text", "conversation_id": "conversation-one", "has_attachments": False, "is_read": False, "categories": []}
    record.update(changes)
    return record


@pytest.fixture
def store(tmp_path):
    result = MailStore(tmp_path / "private" / "mail.sqlite3")
    yield result
    result.close()


def populate(store, records, folder="inbox", *, backend="rest", cursor="cursor-one"):
    store.begin_sync(backend, folder, folder, full=True)
    store.apply_page(backend, folder, folder, records, complete=True, cursor=cursor, seed_done=True)


def test_turkish_fold_is_case_and_ascii_insensitive():
    assert fold("Iıİi Şş Ğğ Üü Öö Çç") == "iiii ss gg uu oo cc"
    assert fold("İŞLEM") == "islem"
    assert fold("go\u0308ru\u0308s\u0327me") == "gorusme"


def test_optimized_fold_preserves_unicode_normalization():
    text = "".join(chr(codepoint) for start, stop in ((0, 0x250), (0x300, 0x370), (0x590, 0x650), (0x1F00, 0x2000), (0xFB00, 0xFB10)) for codepoint in range(start, stop))
    text += "👩🏽‍💻\U0001d185\u034f"
    decomposed = unicodedata.normalize("NFKD", text.casefold().replace("ı", "i"))
    reference = "".join(char for char in decomposed if not unicodedata.combining(char))
    assert fold(text) == reference


@pytest.mark.parametrize("body,query,mode,expected", [
    ("reorganization before organization", "org", "prefix", "[[organization]]"),
    ("organization before organ", "organ", "exact", "[[organ]]"),
    ("Straße görüşmesi", "strass", "prefix", "[[Straße]]"),
    ("Cafe\u0301 go\u0308ru\u0308s\u0327me", "cafe", "exact", "[[Cafe\u0301]]"),
    ("Straße Cafe\u0301 göru\u0308s\u0327me", "gorusme", "prefix", "[[göru\u0308s\u0327me]]"),
])
def test_snippet_highlights_original_word_boundaries_and_unicode(store, body, query, mode, expected):
    populate(store, [message(body=body, subject="", preview="")])
    rows, _ = store.query(query, match_mode=mode)
    assert rows
    assert expected in rows[0]["highlighted"]
    assert "[[reorganization]]" not in rows[0]["highlighted"]
    assert len(rows[0]["snippet"]) <= 242


@pytest.mark.parametrize("query", ["toplantı", "TOPLANTI", "toplan", "sirket", "ŞİRKET", "odeme", "gorus"])
def test_folded_prefix_body_search_and_original_highlight(store, query):
    populate(store, [message()])
    rows, meta = store.query(query, match_mode="prefix")
    assert [row["id"] for row in rows] == ["one"]
    assert "[[" in rows[0]["highlighted"]
    assert "body" not in rows[0]
    assert meta["total_matches"] == 1


def test_stem_and_exact_modes_have_explicit_semantics(store):
    populate(store, [message(body="Toplantı ödeme sigorta")])
    assert store.query("toplantılardan ödemenin sigortalar", match_mode="stem")[0]
    assert not store.query("toplan", match_mode="exact")[0]
    assert store.query('"toplantı ödeme"', match_mode="prefix")[0]
    assert not store.query('"toplantı sigorta"', match_mode="prefix")[0]


def test_literals_do_not_become_fts_operators(store):
    populate(store, [message(body="alpha OR beta")])
    assert store.query("OR", match_mode="exact")[0]
    with pytest.raises(ValueError, match="Unclosed quotation"):
        store.query('"missing close')
    with pytest.raises(ValueError, match="Unknown local query field"):
        store.query("banana:value")
    with pytest.raises(ValueError, match="has:"):
        store.query("has:unread")


@pytest.mark.parametrize("query", ["İstanbul'da", '"İstanbul\'da görüşme"', "'İstanbul görüşmesi'", 'person:"O\'Connor"'])
def test_apostrophes_inside_words_and_quoted_phrases_are_literal(store, query):
    populate(store, [message(body="İstanbul'da görüşme; İstanbul görüşmesi", sender={"name": "O'Connor", "address": "synthetic@example.com"})])
    assert store.query(query)[0]
    with pytest.raises(ValueError, match="Unclosed quotation"):
        store.query("'missing quote")


def test_full_resume_preserves_visible_old_generation(store):
    populate(store, [message("old")])
    original_cursor = store.folder_state("rest", "inbox")["cursor"]
    started = store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="v1")
    store.apply_page("rest", "inbox", "Inbox", [message("new")], next_url="page-two", seed_done=True)
    store.failed("rest", "inbox", "Inbox", "retryable_error")
    assert [row["id"] for row in store.query()[0]] == ["old"]
    state = store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="v1")
    assert state["pending_generation"] == started["pending_generation"]
    assert state["pending_url"] == "page-two"
    assert state["seed_done"] == 1
    assert state["cursor"] == original_cursor
    with pytest.raises(OutlookCliError, match="incomplete"):
        store.query(require_complete=True)
    store.apply_page("rest", "inbox", "Inbox", [message("new-two")], complete=True, cursor="cursor-new")
    assert {row["id"] for row in store.query()[0]} == {"new", "new-two"}
    assert store.folder_state("rest", "inbox")["cursor"] == "cursor-new"


def test_checkpoint_and_page_are_atomic_on_bad_record(store):
    store.begin_sync("rest", "inbox", "Inbox", full=True)
    with pytest.raises(KeyError):
        store.apply_page("rest", "inbox", "Inbox", [message(), {"subject": "missing id"}], next_url="next")
    assert store.db.execute("SELECT count(*) FROM messages").fetchone()[0] == 0
    assert store.folder_state("rest", "inbox")["pending_url"] is None
    store.apply_page("rest", "inbox", "Inbox", [message()], complete=True, cursor="finished")
    assert store.query("toplantı")[0]


def test_existing_sync_ids_look_up_pending_generation_and_chunk_large_batches(store):
    populate(store, [message("old-snapshot")])
    store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="fixture")
    records = [message(f"staged-{index}") for index in range(600)]
    store.apply_page("rest", "inbox", "Inbox", records, next_url="next-page")
    requested = [item["id"] for item in records] + ["old-snapshot", "missing"]
    assert store.existing_sync_ids("rest", "inbox", requested) == {item["id"] for item in records}
    # The durable generation, rather than an in-process set, survives a resume.
    resumed = store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="fixture")
    assert resumed["pending_url"] == "next-page"
    assert store.existing_sync_ids("rest", "inbox", ["staged-1"]) == {"staged-1"}
    assert not store.existing_sync_ids("rest", "inbox", [])


def test_tombstone_ids_survive_resume_but_are_cleared_at_terminal_or_restart(store):
    store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="fixture")
    store.apply_page("rest", "inbox", "Inbox", [], removed=["deleted"], next_url="next-page")
    store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="fixture")
    assert store.existing_sync_ids("rest", "inbox", ["deleted"]) == {"deleted"}
    store.apply_page("rest", "inbox", "Inbox", [], complete=True, cursor="final")
    assert store.db.execute("SELECT count(*) FROM sync_removed").fetchone()[0] == 0
    store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="fixture")
    store.apply_page("rest", "inbox", "Inbox", [], removed=["discard"], next_url="next-page")
    store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="fixture", restart=True)
    assert not store.existing_sync_ids("rest", "inbox", ["discard"])


def test_delta_cursor_advances_only_on_terminal_page(store):
    populate(store, [message("one"), message("two")])
    store.begin_sync("rest", "inbox", "Inbox", full=False)
    store.apply_page("rest", "inbox", "Inbox", [message("one", body="changed")], removed=["two"], next_url="next-delta")
    assert store.read("one")["body"] == "changed"
    assert store.folder_state("rest", "inbox")["cursor"] == "cursor-one"
    assert store.folder_state("rest", "inbox")["pending_url"] == "next-delta"
    assert not store.status()["complete"]
    store.apply_page("rest", "inbox", "Inbox", [], complete=True, cursor="cursor-two")
    assert store.folder_state("rest", "inbox")["cursor"] == "cursor-two"
    assert [row["id"] for row in store.query()[0]] == ["one"]


def test_move_and_old_folder_tombstone_do_not_delete_new_membership(store):
    populate(store, [message("moved")], "inbox")
    populate(store, [message("moved")], "archive")
    assert len(store.query()[0]) == 1
    assert store.query()[0][0]["folder_id"] == "archive"
    store.begin_sync("rest", "inbox", "Inbox", full=False)
    store.apply_page("rest", "inbox", "Inbox", [], removed=["moved"], complete=True, cursor="new")
    assert store.read("moved")["folder_id"] == "archive"


def test_explicit_restart_discards_staging_and_retains_old_snapshot(store):
    populate(store, [message("old")])
    state = store.begin_sync("rest", "inbox", "Inbox", full=True)
    store.apply_page("rest", "inbox", "Inbox", [message("discarded")], next_url="invalid", snapshot_mode=True)
    fresh = store.begin_sync("rest", "inbox", "Inbox", full=True, restart=True)
    assert fresh["pending_generation"] != state["pending_generation"]
    assert fresh["pending_url"] is None
    assert not fresh["snapshot_mode"]
    assert store.db.execute("SELECT count(*) FROM messages WHERE id='discarded'").fetchone()[0] == 0
    assert [row["id"] for row in store.query()[0]] == ["old"]


def test_filters_query_fields_recipients_domains_dates_and_attachments(store):
    attachment = {"id": "att", "name": "ödeme-raporu.pdf", "content_type": "application/pdf", "size": 32, "is_inline": False, "content_bytes": "NEVER STORE"}
    populate(store, [message("match", attachments=[attachment], has_attachments=True),
                     message("other", sender={"name": "Other", "address": "other@other.test"}, to=[], body="irrelevant", received="2026-08-01T00:00:00Z")], "Gelen Kutusu")
    rows, _ = store.query('from:sule@example.com to:gokce@work.test person:"Şule Öztürk" domain:work.test folder:"Gelen Kutusu" after:2026-10-01 before:2026-11-01 has:attachment thread:conversation-one')
    assert [row["id"] for row in rows] == ["match"]
    assert [row["id"] for row in store.query("odeme-rapor")[0]] == ["match"]
    assert "[[ödeme]]-[[raporu]]" in store.query("odeme-rapor")[0][0]["highlighted"]
    assert store.attachments("match") == [{key: attachment[key] for key in ("id", "name", "content_type", "size", "is_inline")}]
    assert "content_bytes" not in store.read("match")["attachments"][0]
    assert [row["id"] for row in store.query(has_attachments=False)[0]] == ["other"]
    assert not store.query(folder="missing")[0]
    assert not store.query(after="2026-10-09T10:00:00Z")[0]
    with pytest.raises(ValueError, match="ISO"):
        store.query(after="not-a-date")


def test_attachment_filter_excludes_inline_only_but_retains_their_metadata(store):
    inline = {"id": "logo", "name": "signature-logo.png", "content_type": "image/png", "size": 42, "is_inline": True}
    ordinary = {"id": "report", "name": "payment-report.pdf", "content_type": "application/pdf", "size": 400, "is_inline": False}
    populate(store, [message("inline", attachments=[inline], has_attachments=False),
                     message("ordinary", attachments=[ordinary], has_attachments=False),
                     message("provider-flag", attachments=[], has_attachments=True)])
    rows, _ = store.query(has_attachments=True)
    assert {row["id"] for row in rows} == {"ordinary", "provider-flag"}
    assert not store.read("inline")["has_attachments"]
    assert store.read("ordinary")["has_attachments"]
    assert store.attachments("inline") == [inline]
    assert store.query("signature-logo")[0][0]["attachment_count"] == 1


def test_compact_repairs_attachment_flags_generated_by_an_older_sync_process(store):
    inline = {"id": "logo", "name": "logo.png", "is_inline": True}
    ordinary = {"id": "report", "name": "report.pdf", "is_inline": False}
    populate(store, [message("inline", attachments=[inline]), message("ordinary", attachments=[ordinary])])
    with store.db:
        store.db.execute("UPDATE messages SET has_attachments=1 WHERE id='inline'")
        store.db.execute("UPDATE messages SET has_attachments=0 WHERE id='ordinary'")
    assert store.compact()["attachment_flags_repaired"] == 2
    assert {row["id"] for row in store.query(has_attachments=True)[0]} == {"ordinary"}
    assert store.compact()["attachment_flags_repaired"] == 0


def test_inventory_does_not_claim_whole_mailbox_from_partial_scope(store):
    store.set_inventory("rest", {"inbox", "archive"})
    populate(store, [message()])
    assert store.status()["complete"]
    assert not store.status()["whole_mailbox_complete"]
    populate(store, [], "archive")
    assert store.status()["whole_mailbox_complete"]
    assert store.status()["scope"] == "whole_mailbox"
    store.begin_sync("rest", "archive", "Archive", full=False)
    assert not store.status()["whole_mailbox_complete"]
    store.apply_page("rest", "archive", "Archive", [], complete=True, cursor="new")
    store.reconcile_folders("rest", {"inbox"})
    assert store.status()["discovered_folder_count"] == 1
    assert store.status()["whole_mailbox_complete"]


def test_compressed_bodies_are_deduplicated_and_html_is_opt_in(store):
    html = '<html><head><style>HIDDEN_STYLE</style></head><body><script>HIDDEN_SCRIPT</script><p>Şirket toplantısı</p>' + '<p>Repeated body content.</p>' * 2000 + '</body></html>'
    populate(store, [message("one", body=html, body_type="HTML"), message("two", body=html, body_type="HTML")])
    assert store.status()["bodies"]["unique"] == 1
    assert store.status()["bodies"]["compressed_bytes"] < len(html.encode()) // 10
    assert "HIDDEN_SCRIPT" not in store.read("one")["body"]
    assert "HIDDEN_STYLE" not in store.read("one")["body"]
    assert store.read("one", body_format="html")["body"] == html
    assert store.read("one")["body_type"] == "Text"
    assert "body" not in store.read("one", body_format="none")
    assert not store.query("HIDDEN_STYLE")[0]
    assert store.query("sirket")[0]
    assert len(store.query("sirket")[0][0]["snippet"]) <= 242
    assert store.db.execute("SELECT body FROM mail_fts LIMIT 1").fetchone()[0] is None


def test_update_delete_fts_old_tokens_are_removed(store):
    populate(store, [message(body="originalword")])
    store.begin_sync("rest", "inbox", "Inbox", full=False)
    store.apply_page("rest", "inbox", "Inbox", [message(body="replacementword")], complete=True, cursor="new")
    assert not store.query("originalword")[0]
    assert store.query("replacementword")[0]
    store.begin_sync("rest", "inbox", "Inbox", full=False)
    store.apply_page("rest", "inbox", "Inbox", [], removed=["one"], complete=True, cursor="newer")
    assert not store.query("replacementword")[0]
    compacted = store.compact()
    assert compacted["file_bytes"].get("-wal", 0) == 0
    assert store.status()["bodies"]["unique"] == 0


def test_offline_research_thread_related_and_body_projection(store):
    populate(store, [message("later"), message("early", received="2026-10-01T00:00:00Z"),
                     message("unrelated-thread", conversation_id="other")])
    thread, meta = store.thread("later", body_format="none")
    assert [record["id"] for record in thread] == ["early", "later"]
    assert [record["id"] for record in store.thread("conversation-one")[0]] == ["early", "later"]
    assert "body" not in thread[0]
    assert not meta["result_complete"]  # No fully discovered mailbox inventory.
    related, meta = store.related("later")
    assert {row["id"] for row in related} == {"early", "unrelated-thread"}
    assert meta["related_person"] == "sule@example.com"
    assert meta["total_matches"] == 2
    assert store.related("later", limit=100000)[0]
    assert "bodies" not in store.query()[1]
    assert "folders" not in store.query()[1]
    with pytest.raises(ResourceNotFoundError):
        store.read("not-local")


@pytest.mark.parametrize("operation", ["read", "thread"])
def test_read_and_thread_keep_one_wal_snapshot_during_promotion(tmp_path, monkeypatch, operation):
    path = tmp_path / "mail.sqlite3"
    writer = MailStore(path)
    populate(writer, [message(body="original body")])
    reader = MailStore(path, readonly=True)
    original_body = reader._body
    changed = False

    def promote_then_read(body_hash, **kwargs):
        nonlocal changed
        if not changed:
            changed = True
            populate(writer, [message(body="replacement body")])
            with writer.db:
                writer.db.execute("DELETE FROM bodies WHERE hash NOT IN (SELECT body_hash FROM messages)")
        return original_body(body_hash, **kwargs)

    monkeypatch.setattr(reader, "_body", promote_then_read)
    try:
        record = reader.read("one") if operation == "read" else reader.thread("one")[0][0]
        assert record["body"] == "original body"
        assert writer.read("one")["body"] == "replacement body"
    finally:
        reader.close()
        writer.close()


def test_attachments_and_coverage_keep_one_snapshot_during_promotion(tmp_path, monkeypatch):
    path = tmp_path / "mail.sqlite3"
    writer = MailStore(path)

    def replace(attachment_id):
        writer.begin_sync("rest", "inbox", "Inbox", full=True, options_signature="text+attachments:v2:rest")
        writer.apply_page("rest", "inbox", "Inbox", [message(attachments=[{"id": attachment_id, "name": attachment_id + ".pdf", "is_inline": False}])], complete=True)

    replace("original")
    reader = MailStore(path, readonly=True)
    find = reader._find

    def promote_after_identity(identity, backend):
        row = find(identity, backend)
        replace("replacement")
        return row

    monkeypatch.setattr(reader, "_find", promote_after_identity)
    try:
        items, meta = reader.attachments("one", include_meta=True)
        assert [item["id"] for item in items] == ["original"]
        assert meta["metadata_complete"]
        assert [item["id"] for item in writer.attachments("one")] == ["replacement"]
    finally:
        reader.close()
        writer.close()


def test_standalone_status_keeps_inventory_and_counts_in_one_snapshot(tmp_path):
    path = tmp_path / "mail.sqlite3"
    writer = MailStore(path)
    populate(writer, [message()])
    writer.set_inventory("rest", {"inbox"})
    reader = MailStore(path, readonly=True)
    original_connection = reader.db

    class PromoteBetweenReads:
        changed = False

        def __getattr__(self, name):
            return getattr(original_connection, name)

        def execute(self, statement, *arguments):
            if statement.startswith("SELECT value FROM meta WHERE key=?") and not self.changed:
                self.changed = True
                populate(writer, [message(), message("two")])
                writer.set_inventory("rest", {"inbox", "archive"})
            return original_connection.execute(statement, *arguments)

    reader.db = PromoteBetweenReads()
    try:
        status = reader.status()
        assert status["message_count"] == 1
        assert status["whole_mailbox_complete"]
        assert writer.status()["message_count"] == 2
        assert not writer.status()["whole_mailbox_complete"]
    finally:
        reader.close()
        writer.close()


@pytest.mark.parametrize("signature,pending,expected", [
    ("legacy:fixture", False, False), ("text+attachments:v1:rest", False, False),
    ("text+attachments:v2:rest", False, True), ("text+attachments:v2:rest", True, False),
])
def test_attachment_metadata_completeness_is_conservative(store, signature, pending, expected):
    store.begin_sync("rest", "inbox", "Inbox", full=True, options_signature=signature)
    store.apply_page("rest", "inbox", "Inbox", [message(has_attachments=True)], complete=True)
    if pending:
        store.begin_sync("rest", "inbox", "Inbox", full=False, options_signature=signature)
    items, meta = store.attachments("one", include_meta=True)
    assert items == []
    assert meta["metadata_complete"] is expected
    assert ("reason" in meta) is not expected


def test_private_files_readonly_connection_and_purge(tmp_path):
    path = tmp_path / "private" / "mail.sqlite3"
    store = MailStore(path)
    populate(store, [message()])
    for entry in (path, Path(str(path) + "-wal"), Path(str(path) + "-shm")):
        if entry.exists():
            assert entry.stat().st_mode & 0o777 == 0o600
    assert path.parent.stat().st_mode & 0o777 == 0o700
    reader = MailStore(path, readonly=True)
    assert reader.query("toplantı")[0]
    with pytest.raises(sqlite3.OperationalError, match="readonly"):
        reader.set_cooldown(1)
    reader.close()
    store.purge()
    assert not path.exists()
    assert not Path(str(path) + "-wal").exists()


def test_symlink_store_refused(tmp_path):
    target = tmp_path / "target"
    target.write_text("retain")
    link = tmp_path / "private" / "mail.sqlite3"
    link.parent.mkdir()
    link.symlink_to(target)
    with pytest.raises(OSError):
        MailStore(link)
    assert target.read_text() == "retain"


@pytest.mark.parametrize("kind", ["symlink", "directory"])
def test_direct_purge_preflights_all_sidecars_before_deleting_store(tmp_path, kind):
    path = tmp_path / "private" / "mail.sqlite3"
    store = MailStore(path)
    populate(store, [message()])
    store.close()
    outside = tmp_path / "outside"
    outside.write_text("retain unrelated data")
    sidecar = Path(str(path) + "-shm")
    if kind == "symlink":
        sidecar.symlink_to(outside)
    else:
        sidecar.mkdir()
    with pytest.raises(OutlookCliError, match="Refusing"):
        store.purge()
    assert path.exists()
    assert outside.read_text() == "retain unrelated data"


def test_persisted_cooldown_does_not_leak_cursor_in_status(store):
    populate(store, [message()], cursor="PRIVATE_CURSOR")
    store.set_cooldown(15)
    assert 0 < store.cooldown_remaining() <= 15
    assert "PRIVATE_CURSOR" not in json.dumps(store.status())
    store.set_cooldown(-1)
    assert store.cooldown_remaining() == 0


def test_legacy_migration_is_sideways_and_retains_source_bytes(tmp_path):
    path = tmp_path / "legacy" / "index.sqlite3"
    legacy = MailIndex(path)
    legacy.apply("rest", "inbox", "Inbox", [message()], full=True, include_body=True)
    legacy.close()
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    store = MailStore(tmp_path / "new" / "mail.sqlite3")
    try:
        outcome = store.migrate_legacy(path)
        assert outcome["source_retained"]
        assert outcome["migrated_messages"] == 1
        assert hashlib.sha256(path.read_bytes()).hexdigest() == before
        assert store.query("sirket")[0]
        assert not store.status()["whole_mailbox_complete"]
        assert store.migrate_legacy(path)["skipped_folders"] == 1
        assert store.read("one")["body"] == message()["body"]
    finally:
        store.close()


def test_preview_only_legacy_migration_does_not_claim_stored_bodies(tmp_path):
    path = tmp_path / "legacy" / "index.sqlite3"
    legacy = MailIndex(path)
    legacy.apply("rest", "inbox", "Inbox", [message(body="", preview="Ödeme toplantısı")], full=True, include_body=False)
    legacy.close()
    store = MailStore(tmp_path / "new" / "mail.sqlite3")
    try:
        store.migrate_legacy(path)
        rows, meta = store.query("odeme")
        assert rows
        assert store.status()["bodies"]["unique"] == 0
        assert "otherwise previews" in meta["body_scope"]
        assert not store.folder_state("rest", "inbox")["include_body"]
        assert store.read("one")["body"] == ""
        assert not meta["whole_mailbox_complete"]
    finally:
        store.close()
