"""Tests for signature_manager.py storage and HTML extraction logic."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from outlook_cli import signature_manager as sm
from outlook_cli.account import AccountPaths
from outlook_cli.exceptions import ResourceNotFoundError


class _Resp:
    def __init__(self, status_code: int = 200, payload: dict | None = None):
        self.status_code = status_code
        self._payload = payload or {}

    def json(self) -> dict:
        return self._payload


def test_list_signatures_returns_sorted_names(monkeypatch, tmp_path):
    paths = AccountPaths("default", tmp_path, tmp_path, tmp_path / "token.json", tmp_path / "browser-state.json", tmp_path / "id_map.json", tmp_path / "scheduled.json", tmp_path, tmp_path / "config.yaml")
    monkeypatch.setattr(sm.account_service, "resolve_account_name", lambda account_name=None: "default")
    monkeypatch.setattr(sm.account_service, "get_account_paths", lambda account_name: paths)
    (tmp_path / "zeta.html").write_text("z")
    (tmp_path / "alpha.html").write_text("a")

    assert sm.list_signatures() == ["alpha", "zeta"]


def test_get_save_and_delete_signature_roundtrip(monkeypatch, tmp_path):
    paths = AccountPaths("default", tmp_path, tmp_path, tmp_path / "token.json", tmp_path / "browser-state.json", tmp_path / "id_map.json", tmp_path / "scheduled.json", tmp_path, tmp_path / "config.yaml")
    monkeypatch.setattr(sm.account_service, "resolve_account_name", lambda account_name=None: "default")
    monkeypatch.setattr(sm.account_service, "get_account_paths", lambda account_name: paths)

    path = sm.save_signature("default", "<b>sig</b>")
    assert path.read_text(encoding="utf-8") == "<b>sig</b>"
    assert sm.get_signature("default") == "<b>sig</b>"

    sm.delete_signature("default")
    assert not path.exists()


def test_get_signature_raises_for_missing_file(monkeypatch, tmp_path):
    paths = AccountPaths("default", tmp_path, tmp_path, tmp_path / "token.json", tmp_path / "browser-state.json", tmp_path / "id_map.json", tmp_path / "scheduled.json", tmp_path, tmp_path / "config.yaml")
    monkeypatch.setattr(sm.account_service, "resolve_account_name", lambda account_name=None: "default")
    monkeypatch.setattr(sm.account_service, "get_account_paths", lambda account_name: paths)

    with pytest.raises(ResourceNotFoundError):
        sm.get_signature("missing")


def test_pull_signature_scans_sent_items_until_it_finds_one(monkeypatch):
    class FakeClient:
        def __init__(self, *args, **kwargs):
            self.calls = 0

        def get(self, *_args, **_kwargs):
            self.calls += 1
            if self.calls == 1:
                return _Resp(payload={"value": [{"Subject": "No sig", "Body": {"Content": "<div>No signature</div>"}}]})
            return _Resp(payload={"value": [{"Subject": "Found sig", "Body": {"Content": "<table><tr><td><a href='mailto:a@b.com'>Mail</a></td></tr></table>"}}]})

        def close(self):
            return None

    monkeypatch.setattr(sm.httpx, "Client", lambda *args, **kwargs: FakeClient())
    monkeypatch.setattr(
        sm,
        "_extract_signature",
        MagicMock(side_effect=[None, "<table><tr><td><a href='mailto:a@b.com'>Mail</a></td></tr></table>"]),
    )

    sig_html, subject = sm.pull_signature("token")

    assert "mailto:" in sig_html
    assert subject == "Found sig"


def test_pull_signature_raises_when_nothing_found(monkeypatch):
    class FakeClient:
        def __init__(self, *args, **kwargs):
            return None

        def get(self, *_args, **_kwargs):
            return _Resp(payload={"value": []})

        def close(self):
            return None

    monkeypatch.setattr(sm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    with pytest.raises(ResourceNotFoundError):
        sm.pull_signature("token")


def test_extract_signature_returns_outermost_matching_table():
    html = """
    <html><body>
      <table id="outer">
        <tr><td>
          <table id="inner"><tr><td><a href="mailto:test@example.com">Mail</a></td></tr></table>
        </td></tr>
      </table>
      <div class="gmail_quote">quoted text</div>
    </body></html>
    """

    extracted = sm._extract_signature(html)

    assert extracted is not None
    assert 'id="outer"' in extracted
    assert "gmail_quote" not in extracted


def test_extract_signature_returns_none_without_mailto():
    assert sm._extract_signature("<table><tr><td>No mail link</td></tr></table>") is None


def test_extract_balanced_table_handles_nested_tables():
    html = "<table><tr><td><table><tr><td>x</td></tr></table></td></tr></table><div>end</div>"

    extracted = sm._extract_balanced_table(html, 0)

    assert extracted == "<table><tr><td><table><tr><td>x</td></tr></table></td></tr></table>"


def test_append_signature_inserts_before_closing_body():
    body, is_html = sm.append_signature("<html><body><p>Hello</p></body></html>", "<b>Sig</b>", True)

    assert is_html is True
    assert "<div id=\"Signature\"><b>Sig</b><br></div></body>" in body


def test_append_signature_converts_plain_text_to_html():
    body, is_html = sm.append_signature("Hello\nWorld", "<b>Sig</b>", False)

    assert is_html is True
    assert body.startswith("<html><body")
    assert "Hello<br>World" in body
    assert "<b>Sig</b>" in body
