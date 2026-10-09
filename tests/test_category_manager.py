"""Tests for category_manager.py OWA calls and bulk category propagation."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import httpx
import pytest

from outlook_cli import category_manager as cm
from outlook_cli.exceptions import (
    PartialFailureError,
    ResourceNotFoundError,
    TokenExpiredError,
)


class _Resp:
    def __init__(self, status_code: int = 200, payload: dict | None = None, headers: dict | None = None):
        self.status_code = status_code
        self._payload = payload or {}
        self.headers = headers or {}
        self.content = b"{}"

    def json(self) -> dict:
        return self._payload

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            request = httpx.Request("POST", "https://example.com")
            response = httpx.Response(self.status_code, request=request)
            raise httpx.HTTPStatusError("failed", request=request, response=response)


def test_owa_request_sends_payload_in_header(monkeypatch):
    post = MagicMock(return_value=_Resp(payload={"ok": True}))
    monkeypatch.setattr(cm.httpx.Client, "request", post)

    result = cm._owa_request("token", "TestAction", {"a": 1})

    assert result == {"ok": True}
    kwargs = post.call_args.kwargs
    assert post.call_args.args[0] == "POST"
    assert kwargs["headers"]["Action"] == "TestAction"
    assert "x-owa-urlpostdata" in kwargs["headers"]
    assert kwargs["content"] == b""


def test_owa_request_raises_for_expired_token(monkeypatch):
    monkeypatch.setattr(cm.httpx.Client, "request", lambda *_args, **_kwargs: _Resp(status_code=401))

    with pytest.raises(TokenExpiredError):
        cm._owa_request("token", "TestAction", {})


def test_update_master_categories_wraps_payload(monkeypatch):
    seen = {}

    def fake_request(token, action, payload):
        seen["token"] = token
        seen["action"] = action
        seen["payload"] = payload
        return {"ok": True}

    monkeypatch.setattr(cm, "_owa_request", fake_request)

    result = cm._update_master_categories("token", add=[{"Name": "Blue"}], remove=["Old"])

    assert result == {"ok": True}
    assert seen["action"] == "UpdateMasterCategoryList"
    assert seen["payload"]["request"]["AddCategoryList"] == [{"Name": "Blue"}]
    assert seen["payload"]["request"]["RemoveCategoryList"] == ["Old"]


def test_get_master_categories_returns_master_list(monkeypatch):
    monkeypatch.setattr(
        cm,
        "_owa_request",
        lambda *_args, **_kwargs: {"MasterCategoryList": {"MasterList": [{"Name": "Finance"}]}},
    )

    assert cm.get_master_categories("token") == [{"Name": "Finance"}]


def test_rename_category_raises_when_category_missing(monkeypatch):
    monkeypatch.setattr(cm, "get_master_categories", lambda _token: [])

    with pytest.raises(ResourceNotFoundError):
        cm.rename_category("token", "Old", "New")


def test_rename_category_can_skip_message_propagation(monkeypatch):
    monkeypatch.setattr(cm, "get_master_categories", lambda _token: [{"Name": "Old", "Id": "1", "Color": 5}])
    update = MagicMock()
    bulk = MagicMock(return_value=99)
    monkeypatch.setattr(cm, "_update_master_categories", update)
    monkeypatch.setattr(cm, "_bulk_rename_on_messages", bulk)

    count = cm.rename_category("token", "Old", "New", propagate=False)

    assert count == 0
    update.assert_called_once()
    bulk.assert_not_called()


def test_bulk_rename_retries_rejection_but_checkpoints_ambiguous_write(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda seconds: sleeps.append(seconds))

    class FakeClient:
        def __init__(self, *args, **kwargs):
            self.get_responses = [
                _Resp(status_code=429, headers={"Retry-After": "1"}),
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old", "Other"]}]}),
                _Resp(payload={"value": []}),
            ]
            self.patch_responses = [httpx.ReadTimeout("slow"), _Resp(payload={})]
            self.get_calls = []
            self.patch_calls = []

        def get(self, url, params=None):
            self.get_calls.append((url, params))
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            self.patch_calls.append((url, json))
            response = self.patch_responses.pop(0)
            if isinstance(response, Exception):
                raise response
            return response

        def request(self, method, url, **kwargs):
            if method == "GET":
                return self.get(url, params=kwargs.get("params"))
            assert method == "PATCH"
            return self.patch(url, json=kwargs.get("json"))

        def close(self):
            return None

    fake_client = FakeClient()
    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: fake_client)

    with pytest.raises(PartialFailureError) as error:
        cm._bulk_rename_on_messages("token", "Old", "New")

    assert Path(error.value.checkpoint).exists()
    assert len(fake_client.patch_calls) == 1
    assert fake_client.patch_calls[-1][1] == {"Categories": ["New", "Other"]}
    assert sleeps == [1]


def test_clear_category_honors_folder_and_max_messages(monkeypatch):
    progress = []

    class FakeClient:
        def __init__(self, *args, **kwargs):
            self.get_calls = []
            self.patch_calls = []

        def get(self, url, params=None):
            self.get_calls.append((url, params))
            return _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old", "Keep"]}, {"Id": "m2", "Categories": ["Old"]}]})

        def patch(self, url, json=None):
            self.patch_calls.append((url, json))
            return _Resp(payload={})

        def request(self, method, url, **kwargs):
            if method == "GET":
                return self.get(url, params=kwargs.get("params"))
            assert method == "PATCH"
            return self.patch(url, json=kwargs.get("json"))

        def close(self):
            return None

    fake_client = FakeClient()
    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: fake_client)

    count = cm.clear_category(
        "token",
        "Old",
        folder="Inbox",
        max_messages=1,
        on_progress=lambda done, total: progress.append((done, total)),
    )

    assert count == 1
    assert fake_client.get_calls[0][0].endswith("/MailFolders/Inbox/messages")
    assert fake_client.patch_calls[0][1] == {"Categories": ["Keep"]}
    assert progress == [(1, -1)]


def test_recolor_category_delegates_to_update(monkeypatch):
    update = MagicMock(return_value={"ok": True})
    monkeypatch.setattr(cm, "_update_master_categories", update)

    result = cm.recolor_category("token", "Finance", 7)

    assert result == {"ok": True}
    update.assert_called_once_with("token", change_color=[{"Name": "Finance", "Color": 7}])
