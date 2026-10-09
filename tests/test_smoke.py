"""Live smoke tests for real Outlook access.

These are intentionally opt-in.

Read-only smoke run:
  OUTLOOK_RUN_SMOKE=1 pytest -m smoke -q

Read-only smoke run with explicit token:
  OUTLOOK_RUN_SMOKE=1 OUTLOOK_TOKEN=... pytest -m smoke -q

Mailbox mutations are tested only with mocks in the unit suite.
"""

from __future__ import annotations

import os
from datetime import datetime, timedelta, timezone

import pytest

from outlook_cli.auth import _load_cached_token, verify_token
from outlook_cli.client import OutlookClient

pytestmark = pytest.mark.smoke


def _require_live_token() -> str:
    if os.environ.get("OUTLOOK_RUN_SMOKE") != "1":
        pytest.skip("Set OUTLOOK_RUN_SMOKE=1 to enable live smoke tests.")
    token = os.environ.get("OUTLOOK_TOKEN") or _load_cached_token()
    if not token:
        pytest.skip("Provide OUTLOOK_TOKEN or log in locally so cached token.json is available.")
    return token


@pytest.fixture
def live_client() -> OutlookClient:
    token = _require_live_token()
    assert verify_token(token) is True
    return OutlookClient(token)


def test_smoke_verify_token():
    token = _require_live_token()
    assert verify_token(token) is True


def test_smoke_whoami(live_client: OutlookClient):
    me = live_client.get_me()
    assert me.get("DisplayName")
    assert me.get("EmailAddress")


def test_smoke_inbox_fetch_top_1(live_client: OutlookClient):
    messages = live_client.get_messages(top=1)

    assert isinstance(messages, list)
    assert len(messages) <= 1
    if messages:
        assert messages[0].id
        assert messages[0].subject


def test_smoke_calendar_fetch_top_1(live_client: OutlookClient):
    now = datetime.now(timezone.utc)
    events = live_client.get_calendar_view(
        start=now.isoformat(),
        end=(now + timedelta(days=30)).isoformat(),
        top=1,
    )

    assert isinstance(events, list)
    assert len(events) <= 1
    if events:
        assert events[0].id
        assert events[0].start


def test_smoke_schedule_list(live_client: OutlookClient):
    entries = live_client.get_scheduled_list()

    assert isinstance(entries, list)
    for entry in entries[:1]:
        assert "subject" in entry
        assert "scheduled_at" in entry


