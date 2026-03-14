"""Tests for _build_query_params() — $filter vs $search logic."""

from __future__ import annotations

from outlook_cli.client import _build_query_params


class TestNoFilters:
    def test_returns_empty(self):
        f, s, needs = _build_query_params()
        assert f == ""
        assert s == ""
        assert needs is False


class TestFilterPath:
    """Pure $filter — supports $orderby."""

    def test_unread_only(self):
        f, s, needs = _build_query_params(unread_only=True)
        assert "IsRead eq false" in f
        assert needs is False

    def test_after_date(self):
        f, s, needs = _build_query_params(filter_after="2026-03-01")
        assert "ReceivedDateTime ge 2026-03-01T00:00:00Z" in f
        assert needs is False

    def test_before_date(self):
        f, s, needs = _build_query_params(filter_before="2026-03-08")
        assert "ReceivedDateTime lt 2026-03-08T23:59:59Z" in f

    def test_category_filter(self):
        f, s, needs = _build_query_params(filter_category="Finance")
        assert "Categories/any(c:c eq 'Finance')" in f
        assert needs is False

    def test_combined_filter(self):
        f, s, needs = _build_query_params(unread_only=True, filter_after="2026-03-01")
        assert "IsRead eq false" in f
        assert "ReceivedDateTime ge" in f
        assert " and " in f
        assert needs is False


class TestSearchPath:
    """KQL $search — can't use $orderby."""

    def test_from_filter(self):
        f, s, needs = _build_query_params(filter_from="alice")
        assert "from:alice" in s
        assert needs is True
        assert f == ""

    def test_subject_filter(self):
        f, s, needs = _build_query_params(filter_subject="Q4 Report")
        assert "subject:Q4 Report" in s
        assert needs is True

    def test_has_attachments(self):
        f, s, needs = _build_query_params(filter_has_attachments=True)
        assert "hasattachments:true" in s
        assert needs is True

    def test_text_plus_date_uses_search(self):
        """When text filters exist, dates go into KQL $search too."""
        f, s, needs = _build_query_params(
            filter_from="bob",
            filter_after="2026-03-01",
        )
        assert needs is True
        assert "from:bob" in s
        assert "received>=2026-03-01" in s
        assert f == ""

    def test_text_plus_category_uses_search(self):
        f, s, needs = _build_query_params(
            filter_subject="invoice",
            filter_category="Finance",
        )
        assert needs is True
        assert "subject:invoice" in s
        assert 'category:"Finance"' in s

    def test_text_plus_unread_uses_search(self):
        f, s, needs = _build_query_params(
            filter_from="alice",
            unread_only=True,
        )
        assert needs is True
        assert "isread:false" in s
        assert "from:alice" in s
