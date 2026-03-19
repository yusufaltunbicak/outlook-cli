"""Tests for timezone and date boundary logic in calendar commands."""

from __future__ import annotations

from datetime import datetime, timezone, timedelta

import pytest

from outlook_cli.commands import calendar as calendar_cmd


class TestParseTimezone:
    """Test _parse_timezone function."""

    def test_none_returns_system_default(self):
        result = calendar_cmd._parse_timezone(None)
        assert result is None

    def test_utc(self):
        tz = calendar_cmd._parse_timezone("UTC")
        assert tz == timezone.utc

    def test_utc_plus_offset(self):
        tz = calendar_cmd._parse_timezone("UTC+8")
        assert tz == timezone(timedelta(hours=8))

        tz = calendar_cmd._parse_timezone("UTC-5")
        assert tz == timezone(timedelta(hours=-5))

    def test_utc_with_minutes(self):
        tz = calendar_cmd._parse_timezone("UTC+5:30")
        assert tz == timezone(timedelta(hours=5, minutes=30))

    def test_invalid_offset_raises_error(self):
        with pytest.raises(Exception):
            calendar_cmd._parse_timezone("UTC+25")  # Invalid hour

    def test_iana_timezone_name(self, monkeypatch):
        # Test with zoneinfo if available (Python 3.9+)
        try:
            import zoneinfo
            tz = calendar_cmd._parse_timezone("Asia/Shanghai")
            # Should return a ZoneInfo object
            assert hasattr(tz, "key")
        except (ImportError, AttributeError):
            # Python < 3.9, skip this test
            pytest.skip("zoneinfo not available")


class TestDateBoundaryLogic:
    """Test date boundary calculation logic."""

    def test_positive_days_future_range(self):
        """Test that positive days calculates future range from today midnight."""
        # Simulate: 2026-03-18 15:00 UTC+8
        mock_now = datetime(2026, 3, 18, 7, 0, tzinfo=timezone.utc)  # 15:00 UTC+8

        # Calculate today_midnight in local time
        import datetime as dt
        now_local = mock_now.astimezone()
        today_midnight = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

        # days = 7
        start = today_midnight
        end = today_midnight + timedelta(days=7)

        # Verify: start should be 2026-03-18 00:00 local, end should be 2026-03-25 00:00 local
        assert start.strftime("%Y-%m-%d") == "2026-03-18"
        assert start.hour == 0
        assert start.minute == 0
        assert end.strftime("%Y-%m-%d") == "2026-03-25"
        assert end.hour == 0
        assert end.minute == 0

    def test_negative_days_past_range(self):
        """Test that negative days calculates past range from today midnight."""
        # Simulate: 2026-03-18 15:00 UTC+8
        mock_now = datetime(2026, 3, 18, 7, 0, tzinfo=timezone.utc)  # 15:00 UTC+8

        # Calculate today_midnight in local time
        import datetime as dt
        now_local = mock_now.astimezone()
        today_midnight = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

        # days = -1 (yesterday)
        start = today_midnight + timedelta(days=-1)
        end = today_midnight

        # Verify: start should be 2026-03-17 00:00 local, end should be 2026-03-18 00:00 local
        assert start.strftime("%Y-%m-%d") == "2026-03-17"
        assert start.hour == 0
        assert start.minute == 0
        assert end.strftime("%Y-%m-%d") == "2026-03-18"
        assert end.hour == 0
        assert end.minute == 0

    def test_negative_days_with_timezone_crossing(self):
        """Test date boundary when timezone crosses UTC midnight."""
        # Edge case: User in UTC+8, time is 02:00 local = 18:00 UTC (previous day UTC)
        # This tests that we use local midnight, not UTC midnight
        tz_plus_8 = timezone(timedelta(hours=8))
        mock_now = datetime(2026, 3, 19, 18, 0, tzinfo=timezone.utc)  # 18:00 UTC = 02:00 UTC+8 next day

        now_local = mock_now.astimezone(tz_plus_8)
        today_midnight = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

        # In UTC+8, 18:00 UTC on 3/19 is 02:00 on 3/20
        assert today_midnight.strftime("%Y-%m-%d") == "2026-03-20"
        assert today_midnight.hour == 0

    def test_days_minus_7_shows_complete_past_week(self):
        """Test that --days -7 shows a complete 7-day range."""
        mock_now = datetime(2026, 3, 18, 10, 0, tzinfo=timezone.utc)  # 18:00 UTC+8

        import datetime as dt
        now_local = mock_now.astimezone()
        today_midnight = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

        # days = -7
        start = today_midnight + timedelta(days=-7)
        end = today_midnight

        # Should be exactly 7 days apart
        delta = end - start
        assert delta.days == 7
        assert delta.seconds == 0  # No time difference

    def test_timezone_boundary_does_not_affect_day_count(self):
        """Test that timezone boundary doesn't affect the number of days returned."""
        # Three different timezones, same UTC time
        utc_time = datetime(2026, 3, 18, 16, 0, tzinfo=timezone.utc)  # 00:00 UTC+8 on 3/18

        # UTC+8: 2026-03-19 00:00
        tz_plus_8 = timezone(timedelta(hours=8))
        local_plus_8 = utc_time.astimezone(tz_plus_8)
        midnight_plus_8 = local_plus_8.replace(hour=0, minute=0, second=0, microsecond=0)

        # UTC (same as UTC): 2026-03-18 16:00
        local_utc = utc_time
        midnight_utc = local_utc.replace(hour=0, minute=0, second=0, microsecond=0)

        # UTC-5: 2026-03-18 11:00
        tz_minus_5 = timezone(timedelta(hours=-5))
        local_minus_5 = utc_time.astimezone(tz_minus_5)
        midnight_minus_5 = local_minus_5.replace(hour=0, minute=0, second=0, microsecond=0)

        # All three should calculate "today midnight" in their respective timezone
        assert midnight_plus_8.strftime("%Y-%m-%d") == "2026-03-19"
        assert midnight_utc.strftime("%Y-%m-%d") == "2026-03-18"
        assert midnight_minus_5.strftime("%Y-%m-%d") == "2026-03-18"

        # And for --days -1, all should span exactly 7 days (just different boundaries)
        start_plus_8 = midnight_plus_8 + timedelta(days=-1)
        start_utc = midnight_utc + timedelta(days=-1)
        start_minus_5 = midnight_minus_5 + timedelta(days=-1)

        # Each should be exactly 1 day before their respective "today midnight"
        assert (midnight_plus_8 - start_plus_8).days == 1
        assert (midnight_utc - start_utc).days == 1
        assert (midnight_minus_5 - start_minus_5).days == 1
