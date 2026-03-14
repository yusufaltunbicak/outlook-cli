"""Tests for pin command — OWA UpdateItem with RenewTime."""

from __future__ import annotations

from unittest.mock import patch

import pytest

from outlook_cli.client import OutlookClient


FAKE_MSG_ID = "AAMk_test_pin_message_id_long_enough_to_pass_resolve_check_1234567"


@pytest.fixture
def client():
    with patch.object(OutlookClient, "_load_id_map", return_value={}):
        c = OutlookClient("fake-token")
    return c


class TestPinMessage:
    def test_pin_calls_owa_update_item(self, client):
        with patch.object(client, "_owa_action", return_value={}) as mock_owa, \
             patch.object(client, "_save_id_map"):
            client.pin_message(FAKE_MSG_ID, pinned=True)

        mock_owa.assert_called_once()
        action = mock_owa.call_args[0][0]
        payload = mock_owa.call_args[0][1]
        assert action == "UpdateItem"
        # Should use SetItemField with RenewTime
        updates = payload["Body"]["ItemChanges"][0]["Updates"]
        assert updates[0]["__type"] == "SetItemField:#Exchange"
        assert updates[0]["Path"]["FieldURI"] == "RenewTime"
        assert updates[0]["Item"]["RenewTime"] == "4500-09-01T00:00:00.000"

    def test_unpin_calls_owa_delete_field(self, client):
        with patch.object(client, "_owa_action", return_value={}) as mock_owa, \
             patch.object(client, "_save_id_map"):
            client.pin_message(FAKE_MSG_ID, pinned=False)

        payload = mock_owa.call_args[0][1]
        updates = payload["Body"]["ItemChanges"][0]["Updates"]
        assert updates[0]["__type"] == "DeleteItemField:#Exchange"
        assert updates[0]["Path"]["FieldURI"] == "RenewTime"

    def test_resolves_display_id(self, client):
        client._id_map["3"] = FAKE_MSG_ID
        with patch.object(client, "_owa_action", return_value={}) as mock_owa, \
             patch.object(client, "_save_id_map"):
            client.pin_message("3", pinned=True)

        payload = mock_owa.call_args[0][1]
        item_id = payload["Body"]["ItemChanges"][0]["ItemId"]["Id"]
        # ID should be converted from URL-safe to standard base64
        expected = FAKE_MSG_ID.replace("-", "/").replace("_", "+")
        assert item_id == expected

    def test_pin_uses_correct_owa_headers(self, client):
        with patch.object(client, "_owa_action", return_value={}) as mock_owa, \
             patch.object(client, "_save_id_map"):
            client.pin_message(FAKE_MSG_ID)

        payload = mock_owa.call_args[0][1]
        assert payload["Header"]["RequestServerVersion"] == "V2018_01_08"
        assert payload["Body"]["ConflictResolution"] == "AlwaysOverwrite"
        assert payload["Body"]["MessageDisposition"] == "SaveOnly"
