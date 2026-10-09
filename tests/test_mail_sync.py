"""Read-only sync protocol regressions with no mailbox or credential access."""
from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace
from urllib.parse import quote

import httpx
import pytest

from outlook_cli import mail_sync
from outlook_cli.constants import BASE_URL
from outlook_cli.exceptions import (
    OutlookCliError,
    RateLimitError,
    ResourceNotFoundError,
)
from outlook_cli.graph import GRAPH_URL
from outlook_cli.mail_sync import (
    GraphSyncReader,
    RestReader,
    removed_id,
    rest_record,
    sync_folder,
)

FOLDER = {"id": "folder-inbox", "displayName": "Inbox"}


def signature(backend="rest", page_size=100):
    return f"text+attachments:v2:{backend}:{page_size}"


def link(name, backend="rest"):
    base = BASE_URL if backend == "rest" else GRAPH_URL
    return f"{base}/mailFolders/folder-inbox/messages/{name}?opaque=fixture-only"


def message(identity, *, parent=FOLDER["id"], subject="Toplantı", backend="rest"):
    if backend == "graph":
        return {"id": identity, "subject": subject, "body": {"content": "Gövde", "contentType": "text"},
                "from": {"emailAddress": {"address": "sender@example.com"}},
                "toRecipients": [], "parentFolderId": parent}
    return {"Id": identity, "Subject": subject, "Body": {"Content": "Gövde", "ContentType": "Text"},
            "From": {"EmailAddress": {"Address": "sender@example.com"}},
            "ToRecipients": [], "ParentFolderId": parent}


def status_error(status):
    request = httpx.Request("GET", "https://fixture.invalid/read-only")
    return httpx.HTTPStatusError("fixture status", request=request,
                                 response=httpx.Response(status, request=request))


class FakeStore:
    """Durable page/checkpoint model; full rounds stage before replacing a snapshot."""
    def __init__(self):
        self.states = {}
        self.active = {}
        self.staging = {}
        self.begins = []
        self.commits = []

    def seed(self, folder_id=FOLDER["id"], *, backend="rest", rows=None, cursor=None):
        key = (backend, folder_id)
        self.active[key] = {row["id"]: deepcopy(row) for row in rows or []}
        self.states[key] = {"cursor": cursor or link("old-delta", backend), "complete": True,
                            "options_signature": signature(backend), "pending_url": None}

    def folder_state(self, backend, identity):
        return deepcopy(self.states.get((backend, identity), {}))

    def begin_sync(self, backend, identity, name, *, full, include_body, options_signature, restart=False):
        key = (backend, identity)
        state = self.states.setdefault(key, {})
        self.begins.append({"full": full, "restart": restart, "identity": identity})
        if restart or not state.get("pending_url"):
            state.update(pending_url=None, seed_done=False, snapshot_mode=False, full=full, pending_full=full)
            if full:
                self.staging[key] = {}
        state.update(options_signature=options_signature, complete=False)
        return deepcopy(state)

    def apply_page(self, backend, identity, name, records, *, removed, next_url, cursor,
                   complete, seed_done, snapshot_mode):
        key = (backend, identity)
        state = self.states[key]
        target = self.staging.setdefault(key, {}) if state.get("full") else self.active.setdefault(key, {})
        for identity_to_remove in removed:
            target.pop(identity_to_remove, None)
        target.update({record["id"]: deepcopy(record) for record in records})
        state.update(pending_url=next_url, complete=complete, seed_done=seed_done, snapshot_mode=snapshot_mode)
        if complete:
            state["cursor"] = cursor
            state["pending_full"] = False
            if state.get("full"):
                self.active[key] = deepcopy(target)
        self.commits.append({"records": deepcopy(records), "removed": list(removed),
                             "complete": complete, "next_url": next_url, "cursor": cursor})


class FakeReader:
    def __init__(self, pages=None, *, backend="rest", hydrated=None):
        self.backend = backend
        self.base_url = BASE_URL if backend == "rest" else GRAPH_URL
        self.page_size = 100
        self.pages = {path: list(responses) for path, responses in (pages or {}).items()}
        self.hydrated = hydrated or {}
        self.calls = []
        self.hydrate_calls = []

    def initial(self, identity):
        if self.backend == "graph":
            path = f"/me/mailFolders/{quote(identity, safe='')}/messages/delta"
        else:
            path = f"/MailFolders/{quote(identity, safe='')}/messages"
        return path, {"$top": self.page_size, "$select": "fixture-fields"}

    def get(self, path, params=None):
        self.calls.append((path, params))
        responses = self.pages.get(path, [])
        assert responses, f"Unexpected fixture read: {path}"
        response = responses.pop(0)
        if isinstance(response, BaseException):
            raise response
        return deepcopy(response)

    def hydrate(self, identity):
        self.hydrate_calls.append(identity)
        response = self.hydrated.get(identity, ResourceNotFoundError("fixture missing"))
        if isinstance(response, BaseException):
            raise response
        return deepcopy(response)

    def record(self, value):
        return {"id": value.get("Id") or value.get("id"),
                "subject": value.get("Subject") or value.get("subject"),
                "body": (value.get("Body") or value.get("body") or {}).get("Content", "")}


def initial_path(backend="rest"):
    return FakeReader(backend=backend).initial(FOLDER["id"])[0]


def test_rest_initial_seed_delta_is_followed_before_completeness():
    store = FakeStore()
    seed, final = link("seed"), link("final-delta")
    reader = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.deltaLink": seed}],
                         seed: [{"value": [], "@odata.deltaLink": final}]})

    result = sync_folder(store, reader, FOLDER)

    assert result["pages"] == 2
    assert result["mode"] == "full"
    assert [commit["complete"] for commit in store.commits] == [False, True]
    assert store.folder_state("rest", FOLDER["id"])["cursor"] == final
    assert list(store.active[("rest", FOLDER["id"])]) == ["a"]
    assert reader.calls[1][1] is None


def test_rest_paginated_seed_delta_is_followed_after_all_seed_pages():
    store = FakeStore()
    page2, seed, final = link("seed-page2"), link("seed-delta"), link("final-delta")
    reader = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.nextLink": page2}],
                         page2: [{"value": [message("b")], "_tracking": True, "@odata.deltaLink": seed}],
                         seed: [{"value": [], "@odata.deltaLink": final}]})

    result = sync_folder(store, reader, FOLDER)

    assert result["pages"] == 3
    assert [commit["complete"] for commit in store.commits] == [False, False, True]
    assert store.folder_state("rest", FOLDER["id"])["cursor"] == final
    assert set(store.active[("rest", FOLDER["id"])]) == {"a", "b"}


def test_interrupted_initial_seed_resumes_saved_delta_without_restarting():
    store = FakeStore()
    seed, final = link("seed"), link("final")
    first = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.deltaLink": seed}]})

    def interrupt(*_args):
        raise KeyboardInterrupt

    with pytest.raises(KeyboardInterrupt):
        sync_folder(store, first, FOLDER, on_page=interrupt)
    assert store.folder_state("rest", FOLDER["id"])["pending_url"] == seed
    assert store.folder_state("rest", FOLDER["id"])["complete"] is False
    second = FakeReader({seed: [{"value": [], "@odata.deltaLink": final}]})

    result = sync_folder(store, second, FOLDER)

    assert result["resumed"] is True
    assert second.calls == [(seed, None)]
    assert list(store.active[("rest", FOLDER["id"])]) == ["a"]


def test_interrupted_paginated_seed_retains_seed_phase_on_resume():
    store = FakeStore()
    page2, seed, final = link("page2"), link("seed"), link("final")
    first = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.nextLink": page2}]})
    with pytest.raises(OutlookCliError, match="page budget"):
        sync_folder(store, first, FOLDER, max_pages=1)
    assert store.folder_state("rest", FOLDER["id"])["seed_done"] is False
    resumed = FakeReader({page2: [{"value": [message("b")], "_tracking": True, "@odata.deltaLink": seed}],
                          seed: [{"value": [], "@odata.deltaLink": final}]})
    result = sync_folder(store, resumed, FOLDER)
    assert result["resumed"] is True
    assert result["mode"] == "full"
    assert result["pages"] == 2
    assert store.folder_state("rest", FOLDER["id"])["cursor"] == final
    assert set(store.active[("rest", FOLDER["id"])]) == {"a", "b"}


def test_incremental_sparse_change_hydrates_current_body_and_metadata():
    store = FakeStore()
    store.seed(rows=[{"id": "a", "subject": "old"}])
    old, final = link("old-delta"), link("new-delta")
    reader = FakeReader({old: [{"value": [{"Id": "a"}], "@odata.deltaLink": final}]},
                        hydrated={"a": message("a", subject="Updated")})

    result = sync_folder(store, reader, FOLDER)

    assert result["mode"] == "delta"
    assert reader.hydrate_calls == ["a"]
    assert store.active[("rest", FOLDER["id"])]["a"]["subject"] == "Updated"


def test_initial_sparse_record_is_hydrated():
    store = FakeStore()
    seed, final = link("seed"), link("final")
    reader = FakeReader({initial_path(): [{"value": [{"Id": "a"}], "_tracking": True, "@odata.deltaLink": seed}],
                         seed: [{"value": [], "@odata.deltaLink": final}]}, hydrated={"a": message("a")})
    sync_folder(store, reader, FOLDER)
    assert reader.hydrate_calls == ["a"]
    assert store.active[("rest", FOLDER["id"])]["a"]["body"] == "Gövde"


@pytest.mark.parametrize("tombstone,identity", [
    ({"id": "Messages('abc')", "@removed": {"reason": "deleted"}}, "abc"),
    ({"id": "messages('a''b')", "reason": "deleted"}, "a'b"),
    ({"id": "graph-id", "@removed": {"reason": "deleted"}}, "graph-id"),
])
def test_removed_id_accepts_legacy_lowercase_and_graph_identity(tombstone, identity):
    assert removed_id(tombstone) == identity


def test_removed_id_rejects_missing_identity():
    with pytest.raises(OutlookCliError, match="no identity"):
        removed_id({"@removed": {"reason": "deleted"}})


@pytest.mark.parametrize("missing", [ResourceNotFoundError("gone"), status_error(404)])
def test_deleted_tombstone_removes_only_selected_folder_scope(missing):
    store = FakeStore()
    store.seed(rows=[{"id": "a", "subject": "old"}])
    store.seed("archive", rows=[{"id": "a", "subject": "archive copy"}])
    reader = FakeReader({link("old-delta"): [{"value": [{"id": "Messages('a')", "@removed": {"reason": "deleted"}}],
                                             "@odata.deltaLink": link("final")}]}, hydrated={"a": missing})

    result = sync_folder(store, reader, FOLDER)

    assert result["removed"] == 1
    assert store.active[("rest", FOLDER["id"])] == {}
    assert "a" in store.active[("rest", "archive")]


def test_older_delete_does_not_erase_message_that_still_exists():
    store = FakeStore()
    store.seed(rows=[{"id": "a", "subject": "old"}])
    reader = FakeReader({link("old-delta"): [{"value": [{"id": "Messages('a')", "@removed": {"reason": "deleted"}}],
                                             "@odata.deltaLink": link("final")}]}, hydrated={"a": message("a", subject="Current")})
    result = sync_folder(store, reader, FOLDER)
    assert result["removed"] == 0
    assert store.active[("rest", FOLDER["id"])]["a"]["subject"] == "Current"


def test_move_out_reconciles_membership_from_current_parent():
    store = FakeStore()
    store.seed(rows=[{"id": "a", "subject": "old"}])
    reader = FakeReader({link("old-delta"): [{"value": [{"Id": "a"}], "@odata.deltaLink": link("final")}]},
                        hydrated={"a": message("a", parent="archive")})
    sync_folder(store, reader, FOLDER)
    assert store.active[("rest", FOLDER["id"])] == {}
    assert store.commits[-1]["removed"] == ["a"]


def test_graph_initial_round_hydrates_attachments_and_uses_final_delta():
    store = FakeStore()
    final = link("final", "graph")
    reader = FakeReader({initial_path("graph"): [{"value": [message("g1", backend="graph")], "@odata.deltaLink": final}]},
                        backend="graph", hydrated={"g1": message("g1", backend="graph")})
    result = sync_folder(store, reader, FOLDER)
    assert result["pages"] == 1
    assert reader.hydrate_calls == ["g1"]
    assert store.folder_state("graph", FOLDER["id"])["cursor"] == final


@pytest.mark.parametrize("bad_link", ["https://other.invalid/collect", "//other.invalid/collect",
                                     BASE_URL + "/messages?token=x#fragment", "https://outlook.office.com/other"])
def test_unsafe_continuation_is_rejected_before_any_page_commit(bad_link):
    store = FakeStore()
    reader = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.nextLink": bad_link}]})
    with pytest.raises(OutlookCliError, match="Refusing"):
        sync_folder(store, reader, FOLDER)
    assert store.commits == []
    assert len(reader.calls) == 1


def test_repeated_continuation_stops_before_recommitting_same_page():
    store = FakeStore()
    repeated = link("repeat")
    reader = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.nextLink": repeated}],
                         repeated: [{"value": [message("b")], "_tracking": True, "@odata.nextLink": repeated}]})
    with pytest.raises(OutlookCliError, match="Repeated"):
        sync_folder(store, reader, FOLDER)
    assert len(store.commits) == 1
    assert store.folder_state("rest", FOLDER["id"])["complete"] is False


def test_expired_delta_reset_retains_previous_complete_snapshot_until_success():
    store = FakeStore()
    store.seed(rows=[{"id": "old", "subject": "Old snapshot"}])
    seed, final = link("seed"), link("new-final")
    reader = FakeReader({link("old-delta"): [status_error(410)],
                         initial_path(): [{"value": [message("new")], "_tracking": True, "@odata.deltaLink": seed}],
                         seed: [{"value": [], "@odata.deltaLink": final}]})
    snapshots = []
    sync_folder(store, reader, FOLDER, on_page=lambda *_: snapshots.append(deepcopy(store.active)))
    assert "old" in snapshots[0][("rest", FOLDER["id"])]
    assert list(store.active[("rest", FOLDER["id"])]) == ["new"]
    assert any(begin["restart"] for begin in store.begins)


def test_expired_pending_page_reset_clears_old_continuation_cycle_history():
    store = FakeStore()
    store.seed(rows=[{"id": "old", "subject": "Old snapshot"}])
    shared, seed, final = link("shared-page"), link("seed"), link("final")
    reader = FakeReader({link("old-delta"): [{"value": [], "@odata.nextLink": shared}],
                         shared: [status_error(410), {"value": [message("new")], "_tracking": True, "@odata.deltaLink": seed}],
                         initial_path(): [{"value": [], "_tracking": True, "@odata.nextLink": shared}],
                         seed: [{"value": [], "@odata.deltaLink": final}]})
    sync_folder(store, reader, FOLDER)
    assert list(store.active[("rest", FOLDER["id"])]) == ["new"]
    assert store.folder_state("rest", FOLDER["id"])["cursor"] == final


def test_repeated_expired_cursor_stops_after_one_reset():
    store = FakeStore()
    store.seed(rows=[{"id": "old", "subject": "Old snapshot"}])
    reader = FakeReader({link("old-delta"): [status_error(410)], initial_path(): [status_error(410)]})
    with pytest.raises(httpx.HTTPStatusError):
        sync_folder(store, reader, FOLDER)
    assert len(reader.calls) == 2
    assert list(store.active[("rest", FOLDER["id"])]) == ["old"]
    assert store.folder_state("rest", FOLDER["id"])["complete"] is False


def test_provider_without_tracking_uses_explicit_snapshot_fallback():
    store = FakeStore()
    reader = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": False}]})
    result = sync_folder(store, reader, FOLDER)
    assert result["mode"] == "snapshot"
    assert store.folder_state("rest", FOLDER["id"])["snapshot_mode"] is True
    assert store.folder_state("rest", FOLDER["id"])["cursor"] is None


def test_missing_delta_checkpoint_retains_old_snapshot_and_stays_incomplete():
    store = FakeStore()
    store.seed(rows=[{"id": "old", "subject": "Old snapshot"}])
    reader = FakeReader({link("old-delta"): [{"value": [message("new")]}]})
    with pytest.raises(OutlookCliError, match="no delta checkpoint"):
        sync_folder(store, reader, FOLDER)
    assert store.commits == []
    assert list(store.active[("rest", FOLDER["id"])]) == ["old"]
    assert store.folder_state("rest", FOLDER["id"])["complete"] is False
    assert store.folder_state("rest", FOLDER["id"])["cursor"] == link("old-delta")


def test_throttling_stops_without_advancing_checkpoint_or_retrying_other_folders():
    store = FakeStore()
    store.seed(rows=[{"id": "old", "subject": "Old snapshot"}])
    reader = FakeReader({link("old-delta"): [RateLimitError("retry budget exhausted")]})
    with pytest.raises(RateLimitError):
        sync_folder(store, reader, FOLDER)
    assert store.commits == []
    assert reader.calls == [(link("old-delta"), None)]
    assert store.folder_state("rest", FOLDER["id"])["complete"] is False


def test_page_budget_checkpoints_for_resume_without_false_completion():
    store = FakeStore()
    next_page = link("page2")
    reader = FakeReader({initial_path(): [{"value": [message("a")], "_tracking": True, "@odata.nextLink": next_page}]})
    with pytest.raises(OutlookCliError, match="page budget"):
        sync_folder(store, reader, FOLDER, max_pages=1)
    state = store.folder_state("rest", FOLDER["id"])
    assert state["pending_url"] == next_page
    assert state["complete"] is False
    assert all(not commit["complete"] for commit in store.commits)


def test_rest_record_retains_recipients_thread_and_attachment_metadata():
    value = message("r1")
    value.update(ConversationId="thread1", BccRecipients=[{"EmailAddress": {"Name": "Bcc", "Address": "bcc@example.com"}}],
                 ReplyTo=[{"EmailAddress": {"Address": "reply@example.com"}}],
                 InternetMessageId="<fixture@example.com>", Attachments=[{"Id": "att", "Name": "report.pdf", "Size": 123,
                                                                        "ContentType": "application/pdf", "IsInline": False}])
    record = rest_record(value)
    assert record["conversation_id"] == "thread1"
    assert record["bcc"][0]["address"] == "bcc@example.com"
    assert record["reply_to"][0]["address"] == "reply@example.com"
    assert record["attachments"][0]["name"] == "report.pdf"
    assert "content_bytes" not in record["attachments"][0]


def test_graph_record_retains_thread_and_attachment_metadata():
    value = message("g1", backend="graph")
    value.update(conversationId="g-thread", bccRecipients=[{"emailAddress": {"address": "bcc@example.com"}}],
                 attachments=[{"id": "att", "name": "report.pdf", "size": 123, "contentType": "application/pdf"}])
    record = GraphSyncReader(None).record(value)
    assert record["conversation_id"] == "g-thread"
    assert record["bcc"][0]["address"] == "bcc@example.com"
    assert record["attachments"][0]["name"] == "report.pdf"


@pytest.mark.parametrize("path,tracking", [
    ("/MailFolders", False),
    ("/MailFolders/folder-id/childFolders", False),
    ("/messages/message-id", False),
    ("/messages('message-id')", False),
    ("/MailFolders/folder-id/messages", True),
    ("/MailFolders/folder-id/messages/", True),
    (BASE_URL + "/MailFolders/folder-id/messages?$skiptoken=fixture-cursor", True),
])
def test_rest_adapter_applies_tracking_only_to_message_collections(monkeypatch, path, tracking):
    captured = []

    def response(client, method, requested_path, **kwargs):
        captured.append((method, requested_path, kwargs))
        return httpx.Response(200, json={"value": []}, headers={"Preference-Applied": "odata.track-changes"} if tracking else {})

    monkeypatch.setattr(mail_sync, "request_response", response)
    client = SimpleNamespace(_client=object(), _refresh=None, account_name="default")
    reader = RestReader(client, interval=0, page_size=50)
    result = reader.get(path, params={"$top": 50})

    method, requested_path, kwargs = captured[0]
    assert method == "GET"
    assert requested_path == path
    prefer = kwargs["headers"]["Prefer"]
    assert ("odata.track-changes" in prefer) is tracking
    assert 'outlook.body-content-type="text"' in prefer
    assert "odata.maxpagesize=50" in prefer
    assert result["_tracking"] is tracking
    assert reader.requests == 1
    assert reader.response_bytes > 0


def test_rest_adapter_rejects_foreign_cursor_before_using_http_client(monkeypatch):
    def forbidden(*_args, **_kwargs):
        raise AssertionError("Foreign origin must be rejected before HTTP")

    monkeypatch.setattr(mail_sync, "request_response", forbidden)
    client = SimpleNamespace(_client=object(), _refresh=None, account_name="default")
    with pytest.raises(OutlookCliError, match="outside the API origin"):
        RestReader(client, interval=0).get("https://foreign.invalid/collect?opaque=cursor")


def test_rest_record_handles_null_draft_sender_recipients_and_body():
    record = rest_record({"Id": "null-draft", "Subject": "Draft", "From": None, "Body": None,
                          "ToRecipients": None, "CcRecipients": None, "Categories": None})
    assert record["id"] == "null-draft"
    assert record["body"] == ""
    assert record["to"] == []


def test_rest_initial_tracking_does_not_cap_whole_folder_with_top():
    reader = RestReader(None, interval=0, page_size=50)
    path, params = reader.initial("folder/encoded")
    assert path == "/MailFolders/folder%2Fencoded/messages"
    assert "$top" not in params
    assert "Body" in params["$select"]
    assert params["$expand"].startswith("Attachments(")


def test_initial_draft_without_sender_does_not_trigger_unnecessary_hydration():
    store = FakeStore()
    seed, final = link("seed"), link("final")
    draft = message("draft")
    draft.pop("From")
    reader = FakeReader({initial_path(): [{"value": [draft], "_tracking": True, "@odata.deltaLink": seed}],
                         seed: [{"value": [], "@odata.deltaLink": final}]})
    sync_folder(store, reader, FOLDER)
    assert reader.hydrate_calls == []
    assert store.active[("rest", FOLDER["id"])]["draft"]["body"] == "Gövde"


def test_initial_missing_body_hydrates_before_committing_page():
    store = FakeStore()
    seed, final = link("seed"), link("final")
    sparse = message("a")
    sparse.pop("Body")
    reader = FakeReader({initial_path(): [{"value": [sparse], "_tracking": True, "@odata.deltaLink": seed}],
                         seed: [{"value": [], "@odata.deltaLink": final}]}, hydrated={"a": message("a")})
    sync_folder(store, reader, FOLDER)
    assert reader.hydrate_calls == ["a"]
    assert store.commits[0]["records"][0]["body"] == "Gövde"
