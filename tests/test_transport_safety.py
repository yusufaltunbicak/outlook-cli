"""Request-scoped retry bounds, throttling delays and credential URL boundaries."""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
from email.utils import format_datetime
from unittest.mock import MagicMock

import httpx
import pytest

from outlook_cli import transport
from outlook_cli.constants import BASE_URL
from outlook_cli.exceptions import (
    AmbiguousMutationError,
    RateLimitError,
    TokenExpiredError,
)


class FakeHTTP:
    def __init__(self, responses):
        self.responses = list(responses)
        self.headers = {"Authorization": "Bearer fixture-secret"}
        self.base_url = BASE_URL
        self.calls = []

    def request(self, method, path, **kwargs):
        self.calls.append((method, path, self.headers.copy(), kwargs))
        outcome = self.responses.pop(0)
        if isinstance(outcome, BaseException):
            raise outcome
        return outcome


def response(status, *, retry_after=None):
    headers = {"Retry-After": retry_after} if retry_after is not None else {}
    return httpx.Response(status, request=httpx.Request("GET", BASE_URL + "/messages"),
                          headers=headers, json={"value": []})


@pytest.mark.parametrize("value,expected", [("0", 0), ("2.5", 2.5), ("-1", None),
                                            ("nan", None), ("inf", None), ("invalid", None), (None, None)])
def test_retry_after_seconds_validates_numeric_delay(value, expected):
    assert transport.retry_after_seconds(value) == expected


def test_retry_after_date_uses_server_deadline_and_clamps_past_dates():
    now = datetime(2026, 10, 9, 12, 0, tzinfo=timezone.utc)
    assert transport.retry_after_seconds(format_datetime(now + timedelta(seconds=120)), now=now) == 120
    assert transport.retry_after_seconds(format_datetime(now - timedelta(seconds=1)), now=now) == 0


def test_throttle_beyond_retry_budget_returns_required_cooldown_without_sleep(monkeypatch, capsys):
    client = FakeHTTP([response(429, retry_after="120")])
    sleep = MagicMock()
    monkeypatch.setattr(transport.time, "sleep", sleep)

    with pytest.raises(RateLimitError) as error:
        transport.request_response(client, "GET", "/messages", deadline=90)

    assert error.value.retry_after == 120
    assert len(client.calls) == 1
    sleep.assert_not_called()
    assert "fixture-secret" not in str(error.value)
    assert capsys.readouterr().out == ""


def test_throttle_http_date_is_retained_for_next_run_cooldown(monkeypatch):
    now = datetime(2026, 10, 9, 12, 0, tzinfo=timezone.utc)
    parse = transport.retry_after_seconds
    monkeypatch.setattr(transport, "retry_after_seconds", lambda value: parse(value, now=now))
    client = FakeHTTP([response(429, retry_after=format_datetime(now + timedelta(seconds=180)))])
    with pytest.raises(RateLimitError) as error:
        transport.request_response(client, "GET", "/messages", deadline=90)
    assert error.value.retry_after == 180
    assert len(client.calls) == 1


def test_rejected_read_honors_retry_after_then_stops_at_retry_count(monkeypatch):
    client = FakeHTTP([response(429, retry_after="2"), response(429, retry_after="9")])
    sleeps = []
    monkeypatch.setattr(transport.time, "sleep", sleeps.append)
    with pytest.raises(RateLimitError) as error:
        transport.request_response(client, "GET", "/messages", max_retries=1)
    assert sleeps == [2]
    assert error.value.retry_after == 9
    assert len(client.calls) == 2


def test_auth_refresh_retries_only_the_rejected_http_request():
    client = FakeHTTP([response(401), response(200)])
    refresh = MagicMock(return_value="fresh-fixture-secret")
    transport.request_response(client, "GET", "/messages/fixture", refresh=refresh)
    assert [call[:2] for call in client.calls] == [("GET", "/messages/fixture"), ("GET", "/messages/fixture")]
    assert client.calls[0][2]["Authorization"] == "Bearer fixture-secret"
    assert client.calls[1][2]["Authorization"] == "Bearer fresh-fixture-secret"
    refresh.assert_called_once()


def test_repeated_401_does_not_refresh_forever():
    client = FakeHTTP([response(401), response(401)])
    refresh = MagicMock(return_value="fresh-fixture-secret")
    with pytest.raises(TokenExpiredError):
        transport.request_response(client, "GET", "/messages/fixture", refresh=refresh)
    assert len(client.calls) == 2
    refresh.assert_called_once()


@pytest.mark.parametrize("outcome", [httpx.ReadTimeout("fixture network error"), response(503)])
def test_uncertain_write_is_not_replayed(outcome):
    client = FakeHTTP([outcome])
    with pytest.raises(AmbiguousMutationError):
        transport.request_response(client, "POST", "/messages", json={"fixture": True})
    assert len(client.calls) == 1


@pytest.mark.parametrize("url", ["https://foreign.invalid/collect", "http://outlook.office.com/api/v2.0/me/messages",
                                  "https://outlook.office.com/other", "https://user:password@outlook.office.com/api/v2.0/me/messages"])
def test_pagination_never_accepts_a_url_that_could_expose_bearer_credentials(url):
    with pytest.raises(ValueError, match="Refusing pagination"):
        transport.ensure_next_link(FakeHTTP([]), url)
