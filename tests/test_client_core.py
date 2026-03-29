"""Tests for OutlookClient core HTTP, scheduling, and calendar behavior."""

from __future__ import annotations

from unittest.mock import MagicMock

import httpx
import pytest

from outlook_cli.client import OutlookClient
from outlook_cli.exceptions import RateLimitError, ResourceNotFoundError, TokenExpiredError


class _Resp:
    def __init__(self, status_code: int = 200, payload: dict | None = None, headers: dict | None = None, content: bytes | None = None):
        self.status_code = status_code
        self._payload = payload or {}
        self.headers = headers or {}
        if content is not None:
            self.content = content
        elif status_code == 204:
            self.content = b""
        else:
            self.content = b"{}"

    def json(self) -> dict:
        return self._payload

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            request = httpx.Request("GET", "https://example.com")
            response = httpx.Response(self.status_code, request=request)
            raise httpx.HTTPStatusError("failed", request=request, response=response)


@pytest.fixture
def client(monkeypatch):
    monkeypatch.setattr(OutlookClient, "_load_id_map", lambda self: {})
    return OutlookClient("fake-token")


def test_request_raises_token_expired_on_401(client, monkeypatch):
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: _Resp(status_code=401))

    with pytest.raises(TokenExpiredError):
        client._request("GET", "/messages")


def test_request_retries_on_429_then_succeeds(client, monkeypatch):
    responses = iter([
        _Resp(status_code=429, headers={"Retry-After": "1"}),
        _Resp(payload={"value": [1]}),
    ])
    sleeps = []
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: next(responses))
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda seconds: sleeps.append(seconds))

    result = client._request("GET", "/messages")

    assert result == {"value": [1]}
    assert sleeps == [1]


def test_request_raises_rate_limit_after_three_retries(client, monkeypatch):
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: _Resp(status_code=429, headers={"Retry-After": "1"}))
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda *_args, **_kwargs: None)

    with pytest.raises(RateLimitError):
        client._request("GET", "/messages")


def test_request_returns_empty_dict_on_204(client, monkeypatch):
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: _Resp(status_code=204))

    assert client._request("DELETE", "/messages/1") == {}


def test_resolve_id_uses_map_or_passes_through(client):
    client._id_map["3"] = "real-id"

    assert client._resolve_id("3") == "real-id"
    assert client._resolve_id("x" * 60) == "x" * 60


def test_resolve_id_raises_for_unknown_display_number(client):
    with pytest.raises(ResourceNotFoundError):
        client._resolve_id("99")


def test_get_open_target_prefers_message_link(client, monkeypatch):
    client._id_map["3"] = "msg-id"
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"WebLink": "https://outlook.office365.com/owa/?ItemID=msg-id"}

    monkeypatch.setattr(client, "_get", fake_get)

    kind, url = client.get_open_target("3")

    assert kind == "message"
    assert url == "https://outlook.office365.com/owa/?ItemID=msg-id"
    assert calls == [("/messages/msg-id", {"$select": "WebLink"})]


def test_get_open_target_falls_back_to_event_link(client, monkeypatch):
    client._id_map["42"] = "event-id"
    request = httpx.Request("GET", "https://example.com")
    not_found = httpx.Response(404, request=request)

    def fake_get(path, params=None):
        if path == "/messages/event-id":
            raise httpx.HTTPStatusError("not found", request=request, response=not_found)
        if path == "/events/event-id":
            return {"WebLink": "https://outlook.office365.com/owa/?itemid=event-id"}
        raise AssertionError(f"unexpected path: {path}")

    monkeypatch.setattr(client, "_get", fake_get)

    kind, url = client.get_open_target("42")

    assert kind == "event"
    assert url == "https://outlook.office365.com/owa/?itemid=event-id"


def test_get_open_target_raises_generic_missing_item_error(client):
    with pytest.raises(ResourceNotFoundError, match="Unknown item #99"):
        client.get_open_target("99")


def test_assign_display_nums_reuses_existing_and_evicts_old_entries(client, monkeypatch, make_email):
    client.MAX_ID_MAP_SIZE = 2
    client._id_map = {"1": "existing-id"}
    client._next_num = 2
    monkeypatch.setattr(client, "_save_id_map", lambda: None)

    messages = [
        make_email(id="existing-id"),
        make_email(id="new-1"),
        make_email(id="new-2"),
    ]

    client._assign_display_nums(messages)

    assert messages[0].display_num == 1
    assert messages[1].display_num == 2
    assert messages[2].display_num == 3
    assert "1" not in client._id_map
    assert client._id_map == {"2": "new-1", "3": "new-2"}


def test_get_messages_with_no_category_overfetches_until_top(client, monkeypatch):
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    responses = iter([
        {
            "value": [
                {"Id": "1", "Categories": ["Finance"]},
                {"Id": "2", "Categories": []},
                {"Id": "3", "Categories": ["Urgent"]},
                {"Id": "4", "Categories": []},
                {"Id": "5", "Categories": ["Other"]},
                {"Id": "6", "Categories": ["More"]},
            ]
        }
    ])
    monkeypatch.setattr(client, "_get", lambda *_args, **_kwargs: next(responses))

    messages = client.get_messages(top=2, filter_no_category=True)

    assert [m.id for m in messages] == ["2", "4"]
    assert all(not m.categories for m in messages)


def test_get_messages_resolves_named_folder_for_standard_listing(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "folder-id")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_messages(folder="Onaylar", top=5)

    assert calls == [
        (
            "/MailFolders/folder-id/messages",
            {"$top": 5, "$skip": 0, "$orderby": "ReceivedDateTime desc"},
        )
    ]


def test_get_messages_uses_folder_scoped_search_for_text_filters(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "folder-id")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_messages(folder="Onaylar", top=5, filter_subject="Tüsaf")

    assert calls == [
        (
            "/MailFolders/folder-id/messages",
            {"$top": 5, "$search": '"subject:Tüsaf"'},
        )
    ]


def test_schedule_send_tracks_entry(client, monkeypatch):
    send_mail = MagicMock()
    track = MagicMock(return_value={"subject": "Planned"})
    monkeypatch.setattr(client, "send_mail", send_mail)
    monkeypatch.setattr(client, "_track_scheduled", track)

    result = client.schedule_send(["a@example.com"], "Planned", "Body", "2026-03-20T10:00:00Z")

    assert result == {"subject": "Planned"}
    send_mail.assert_called_once()
    track.assert_called_once_with(to=["a@example.com"], cc=None, subject="Planned", send_at="2026-03-20T10:00:00Z")


def test_schedule_draft_patches_sends_and_tracks(client, monkeypatch):
    client._id_map["7"] = "draft-id"
    monkeypatch.setattr(
        client,
        "_get",
        lambda *_args, **_kwargs: {
            "Subject": "Draft subject",
            "ToRecipients": [{"EmailAddress": {"Address": "a@example.com"}}],
            "CcRecipients": [{"EmailAddress": {"Address": "b@example.com"}}],
        },
    )
    patch = MagicMock(return_value={"Id": "updated-id"})
    post = MagicMock(return_value={})
    track = MagicMock(return_value={"message_id": "updated-id"})
    monkeypatch.setattr(client, "_patch", patch)
    monkeypatch.setattr(client, "_post", post)
    monkeypatch.setattr(client, "_track_scheduled", track)

    result = client.schedule_draft("7", "2026-03-20T10:00:00Z")

    assert result == {"message_id": "updated-id"}
    patch.assert_called_once()
    post.assert_called_once_with("/messages/updated-id/send")
    track.assert_called_once_with(
        to=["a@example.com"],
        cc=["b@example.com"],
        subject="Draft subject",
        send_at="2026-03-20T10:00:00Z",
        message_id="updated-id",
    )


def test_get_scheduled_list_enriches_entries_with_draft_ids(client, monkeypatch):
    monkeypatch.setattr(client, "_load_scheduled", lambda: [{"subject": "Draft A", "scheduled_at": "x"}])
    monkeypatch.setattr(
        client,
        "_get",
        lambda *_args, **_kwargs: {"value": [{"Id": "draft-1", "Subject": "Draft A"}]},
    )

    entries = client.get_scheduled_list()

    assert entries[0]["message_id"] == "draft-1"


def test_cancel_scheduled_entry_removes_local_and_server_copy(client, monkeypatch):
    local_entries = [{"subject": "A", "scheduled_at": "x"}]
    monkeypatch.setattr(client, "get_scheduled_list", lambda: [{"subject": "A", "scheduled_at": "x", "message_id": "draft-1"}])
    monkeypatch.setattr(client, "_load_scheduled", lambda: list(local_entries))
    saved = {}
    monkeypatch.setattr(client, "_save_scheduled", lambda entries: saved.setdefault("entries", entries))
    delete = MagicMock()
    monkeypatch.setattr(client, "_delete", delete)

    removed = client.cancel_scheduled_entry(1)

    assert removed["server_deleted"] is True
    assert saved["entries"] == []
    delete.assert_called_once_with("/messages/draft-1")


def test_copy_message_posts_to_copy_action_with_resolved_folder(client, monkeypatch):
    client._id_map["3"] = "real-msg-id"
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: f"resolved:{folder}")
    post = MagicMock(return_value={"Id": "copied-msg-id", "Subject": "Copied"})
    monkeypatch.setattr(client, "_post", post)

    email = client.copy_message("3", "Archive")

    assert email.id == "copied-msg-id"
    post.assert_called_once_with("/messages/real-msg-id/copy", json={"DestinationId": "resolved:Archive"})


def test_create_event_builds_expected_payload(client, monkeypatch):
    captured = {}

    def fake_post(path, json=None):
        captured["path"] = path
        captured["json"] = json
        return {
            "Id": "ev-1",
            "Subject": "Planning",
            "Start": {"DateTime": "2026-03-20T10:00:00"},
            "End": {"DateTime": "2026-03-20T11:00:00"},
        }

    monkeypatch.setattr(client, "_post", fake_post)

    event = client.create_event(
        "Planning",
        "2026-03-20T10:00:00",
        "2026-03-20T11:00:00",
        timezone="Europe/Istanbul",
        attendees=["a@example.com"],
        location="Board Room",
        body="<b>Hello</b>",
        html=True,
        is_all_day=False,
        reminder_minutes=30,
        is_online_meeting=True,
        recurrence={"Pattern": {"Type": "Daily"}, "Range": {"Type": "Numbered"}},
    )

    assert event.subject == "Planning"
    assert captured["path"] == "/events"
    assert captured["json"]["Attendees"][0]["EmailAddress"]["Address"] == "a@example.com"
    assert captured["json"]["Location"]["DisplayName"] == "Board Room"
    assert captured["json"]["Body"]["ContentType"] == "HTML"
    assert captured["json"]["OnlineMeetingProvider"] == "TeamsForBusiness"


def test_event_attendee_helpers_merge_and_filter_existing(client, monkeypatch):
    client._id_map["5"] = "ev-5"
    current = {
        "Attendees": [
            {"EmailAddress": {"Address": "a@example.com"}, "Type": "Required"},
            {"EmailAddress": {"Address": "b@example.com"}, "Type": "Required"},
        ]
    }
    monkeypatch.setattr(client, "_get", lambda *_args, **_kwargs: current)
    patch = MagicMock(
        return_value={
            "Id": "ev-5",
            "Subject": "Standup",
            "Start": {"DateTime": "2026-03-20T10:00:00"},
            "End": {"DateTime": "2026-03-20T11:00:00"},
        }
    )
    monkeypatch.setattr(client, "_patch", patch)

    client.add_event_attendees("5", ["a@example.com", "c@example.com"])
    add_payload = patch.call_args.kwargs["json"]
    assert [a["EmailAddress"]["Address"] for a in add_payload["Attendees"]] == [
        "a@example.com",
        "b@example.com",
        "c@example.com",
    ]

    client.remove_event_attendees("5", ["b@example.com"])
    remove_payload = patch.call_args.kwargs["json"]
    assert [a["EmailAddress"]["Address"] for a in remove_payload["Attendees"]] == [
        "a@example.com",
        "c@example.com",
    ]


def test_get_event_instances_uses_series_master_id(client, monkeypatch):
    client._id_map["3"] = "occurrence-id"
    monkeypatch.setattr(client, "_assign_event_display_nums", lambda events: None)
    responses = iter([
        {"Type": "Occurrence", "SeriesMasterId": "master-id"},
        {"value": [{"Id": "inst-1", "Start": {"DateTime": "2026-03-21T10:00:00"}, "End": {"DateTime": "2026-03-21T11:00:00"}}]},
    ])
    calls = []

    def fake_get(path, params=None):
        calls.append(path)
        return next(responses)

    monkeypatch.setattr(client, "_get", fake_get)

    events = client.get_event_instances("3", "2026-03-20T00:00:00Z", "2026-03-25T00:00:00Z")

    assert len(events) == 1
    assert calls[1] == "/events/master-id/instances"


def test_respond_to_event_posts_expected_payload(client, monkeypatch):
    client._id_map["2"] = "event-id"
    post = MagicMock(return_value={})
    monkeypatch.setattr(client, "_post", post)

    client.respond_to_event("2", "accept", comment="Works for me", send_response=False)

    post.assert_called_once_with("/events/event-id/accept", json={"SendResponse": False, "Comment": "Works for me"})


def test_resolve_calendar_supports_exact_and_partial_match(client, monkeypatch):
    monkeypatch.setattr(
        client,
        "get_calendars",
        lambda: [{"Id": "1", "Name": "Primary"}, {"Id": "2", "Name": "Team Calendar"}],
    )

    assert client._resolve_calendar("Primary") == "1"
    assert client._resolve_calendar("Team") == "2"


def test_resolve_calendar_raises_for_missing_name(client, monkeypatch):
    monkeypatch.setattr(client, "get_calendars", lambda: [{"Id": "1", "Name": "Primary"}])

    with pytest.raises(ResourceNotFoundError):
        client._resolve_calendar("Missing")


def test_get_master_categories_calls_owa_action(client, monkeypatch):
    owa = MagicMock(return_value={"Body": {"CategoryDetailsList": []}})
    monkeypatch.setattr(client, "_owa_action", owa)

    client.get_master_categories()

    assert owa.call_args.args[0] == "FindCategoryDetails"


# ── Plain text to HTML auto-conversion ───────────────────


def test_plain_text_to_html_preserves_line_breaks():
    from outlook_cli.client import _plain_text_to_html

    result = _plain_text_to_html("Hello\nWorld")
    assert result == "Hello<br>\nWorld"


def test_plain_text_to_html_escapes_html_chars():
    from outlook_cli.client import _plain_text_to_html

    result = _plain_text_to_html("A < B & C > D")
    assert "&lt;" in result
    assert "&amp;" in result
    assert "&gt;" in result
    assert "<br>" not in result  # no newlines = no <br>


def test_send_mail_auto_converts_plain_text(client, monkeypatch):
    captured = {}
    monkeypatch.setattr(client, "_post", lambda path, json=None: captured.update(json=json))

    client.send_mail(to=["a@b.com"], subject="Test", body="Line1\nLine2", html=False)

    body = captured["json"]["Message"]["Body"]
    assert body["ContentType"] == "HTML"
    assert "Line1<br>\nLine2" in body["Content"]


def test_send_mail_passes_html_body_unchanged(client, monkeypatch):
    captured = {}
    monkeypatch.setattr(client, "_post", lambda path, json=None: captured.update(json=json))

    client.send_mail(to=["a@b.com"], subject="Test", body="<b>Bold</b>", html=True)

    body = captured["json"]["Message"]["Body"]
    assert body["ContentType"] == "HTML"
    assert body["Content"] == "<b>Bold</b>"


def test_create_draft_auto_converts_plain_text(client, monkeypatch):
    monkeypatch.setattr(
        client, "_post", lambda path, json=None: {"Id": "d1", "Subject": "S"}
    )
    monkeypatch.setattr(client, "_save_id_map", lambda: None)

    email = client.create_draft(to=["a@b.com"], subject="S", body="A\nB", html=False)

    assert email.id == "d1"
