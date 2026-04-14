"""
Tests for apps/labor_parts_analyzer/analyze.py
Run from repo root:  python -m pytest apps/labor_parts_analyzer/tests/test_analyze.py -v
"""

import sys
import csv
import io
from datetime import datetime
from pathlib import Path

import pytest

# Allow imports from apps/labor_parts_analyzer
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from analyze import (
    classify_description,
    parse_ro_open_date,
    parse_timestamp,
    ro_age_hours,
    classify_row,
    run_analysis,
    validate_headers,
    WINDOW_PENDING,
    WINDOW_HARD,
    WINDOW_DIVERGE,
    WINDOW_REVIEW,
    PENDING_HOURS,
    HARD_MAX_DAYS,
    DIVERGE_MIN_DAYS,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

KEYWORDS = ["replace", "install", "new", "kit", "assembly"]
NEGATORS = ["inspect", "check", "verify", "test", "adjust", "measure", "clean", "diagnose", "drain"]

NOW = datetime(2025, 11, 10, 12, 0, 0)  # fixed reference point


def _row(desc="replace oil filter", parts_found="False",
         ro_open_date="", timestamp="11/10/2025 12:00:00 PM",
         ro_number="123456", labor_id="L 1", sequence="101"):
    return {
        "Timestamp":    timestamp,
        "RO_Number":    ro_number,
        "Labor_ID":     labor_id,
        "Description":  desc,
        "Parts_Found":  parts_found,
        "Sequence":     sequence,
        "RO_Open_Date": ro_open_date,
    }


# ---------------------------------------------------------------------------
# classify_description
# ---------------------------------------------------------------------------

class TestClassifyDescription:
    def test_keyword_match_no_negator(self):
        assert classify_description("replace oil filter", KEYWORDS, NEGATORS) is True

    def test_keyword_with_negator_returns_false(self):
        assert classify_description("inspect and replace", KEYWORDS, NEGATORS) is False

    def test_negator_only_returns_false(self):
        assert classify_description("check tire pressure", KEYWORDS, NEGATORS) is False

    def test_no_keyword_no_negator_returns_false(self):
        assert classify_description("rotate tires", KEYWORDS, NEGATORS) is False

    def test_case_insensitive_keyword(self):
        assert classify_description("REPLACE BRAKE PADS", KEYWORDS, NEGATORS) is True

    def test_case_insensitive_negator(self):
        assert classify_description("INSPECT brake pads", KEYWORDS, NEGATORS) is False

    def test_install_keyword(self):
        assert classify_description("install new wiper blades", KEYWORDS, NEGATORS) is True

    def test_verify_negator(self):
        assert classify_description("verify oil level new", KEYWORDS, NEGATORS) is False

    def test_empty_description(self):
        assert classify_description("", KEYWORDS, NEGATORS) is False

    def test_substring_match(self):
        # "newly" contains "new" as substring
        assert classify_description("newly installed part", KEYWORDS, NEGATORS) is True


# ---------------------------------------------------------------------------
# parse_ro_open_date
# ---------------------------------------------------------------------------

class TestParseRoOpenDate:
    def test_valid_ddmmmyy(self):
        result = parse_ro_open_date("05NOV25", NOW)
        assert result == datetime(2025, 11, 5, 0, 0, 0)

    def test_valid_lowercase(self):
        result = parse_ro_open_date("05nov25", NOW)
        assert result == datetime(2025, 11, 5, 0, 0, 0)

    def test_unknown_string_returns_none(self):
        assert parse_ro_open_date("UNKNOWN", NOW) is None

    def test_empty_string_returns_none(self):
        assert parse_ro_open_date("", NOW) is None

    def test_invalid_format_returns_none(self):
        assert parse_ro_open_date("2025-11-05", NOW) is None

    def test_another_valid_date(self):
        result = parse_ro_open_date("01JAN25", NOW)
        assert result == datetime(2025, 1, 1, 0, 0, 0)


# ---------------------------------------------------------------------------
# ro_age_hours
# ---------------------------------------------------------------------------

class TestRoAgeHours:
    def test_uses_ro_open_date_when_available(self):
        # 05NOV25 → 5 days before NOW (10NOV25 12:00) = 5*24 + 12 = 132 hours
        row = _row(ro_open_date="05NOV25", timestamp="11/01/2025 12:00:00 PM")
        age = ro_age_hours(row, NOW)
        assert age == pytest.approx(5 * 24 + 12, abs=0.1)

    def test_falls_back_to_timestamp_when_ro_open_date_unknown(self):
        # timestamp 5 days before NOW
        row = _row(ro_open_date="UNKNOWN", timestamp="11/05/2025 12:00:00 PM")
        age = ro_age_hours(row, NOW)
        assert age == pytest.approx(5 * 24, abs=0.1)

    def test_returns_none_when_both_unparseable(self):
        row = _row(ro_open_date="UNKNOWN", timestamp="bad-timestamp")
        assert ro_age_hours(row, NOW) is None

    def test_zero_age(self):
        row = _row(ro_open_date="", timestamp="11/10/2025 12:00:00 PM")
        age = ro_age_hours(row, NOW)
        assert age == pytest.approx(0.0, abs=0.1)


# ---------------------------------------------------------------------------
# classify_row — window boundaries
# ---------------------------------------------------------------------------

class TestClassifyRowWindows:
    def _make_now_minus_hours(self, h):
        from datetime import timedelta
        return (NOW - timedelta(hours=h)).strftime("%m/%d/%Y %I:%M:%S %p")

    def test_pending_lead_time_under_48h(self):
        row = _row(timestamp=self._make_now_minus_hours(24), desc="replace filter")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_PENDING

    def test_exactly_48h_is_still_pending(self):
        row = _row(timestamp=self._make_now_minus_hours(48), desc="replace filter")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_PENDING

    def test_hard_block_49h_keyword_no_parts(self):
        row = _row(timestamp=self._make_now_minus_hours(49),
                   desc="replace filter", parts_found="False")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_HARD

    def test_hard_block_requires_keyword(self):
        row = _row(timestamp=self._make_now_minus_hours(72),
                   desc="rotate tires", parts_found="False")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_REVIEW

    def test_hard_block_parts_found_true_goes_to_review(self):
        row = _row(timestamp=self._make_now_minus_hours(72),
                   desc="replace filter", parts_found="True")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_REVIEW

    def test_exactly_7_days_is_hard_block(self):
        row = _row(timestamp=self._make_now_minus_hours(7 * 24),
                   desc="install kit", parts_found="False")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_HARD

    def test_divergence_at_8_days(self):
        row = _row(timestamp=self._make_now_minus_hours(DIVERGE_MIN_DAYS * 24 + 1))
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_DIVERGE

    def test_divergence_at_10_days(self):
        row = _row(timestamp=self._make_now_minus_hours(240))
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_DIVERGE

    def test_unparseable_timestamp_goes_to_review(self):
        row = _row(timestamp="not-a-date", ro_open_date="UNKNOWN")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_REVIEW

    def test_ro_open_date_takes_priority_over_timestamp(self):
        # ro_open_date says 1 day ago (pending) but timestamp says 10 days ago
        from datetime import timedelta
        recent_open  = (NOW - timedelta(hours=12)).strftime("%d%b%y").upper()
        old_timestamp = (NOW - timedelta(days=10)).strftime("%m/%d/%Y %I:%M:%S %p")
        row = _row(ro_open_date=recent_open, timestamp=old_timestamp, desc="replace filter")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_PENDING

    def test_negator_overrides_keyword_in_hard_block_window(self):
        # "inspect and replace" — negator present, should not be HARD_BLOCK
        row = _row(timestamp=self._make_now_minus_hours(72),
                   desc="inspect and replace filter", parts_found="False")
        assert classify_row(row, KEYWORDS, NEGATORS, NOW) == WINDOW_REVIEW


# ---------------------------------------------------------------------------
# run_analysis (integration)
# ---------------------------------------------------------------------------

class TestRunAnalysis:
    def _make_csv(self, rows: list[dict], tmp_path: Path) -> Path:
        p = tmp_path / "Master_Labor_Log.csv"
        headers = list(rows[0].keys())
        with open(p, "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=headers)
            w.writeheader()
            w.writerows(rows)
        return p

    def _row_ts(self, hours_ago: int, desc="replace filter",
                parts_found="False", **kwargs):
        from datetime import timedelta
        ts = (NOW - timedelta(hours=hours_ago)).strftime("%m/%d/%Y %I:%M:%S %p")
        return _row(desc=desc, parts_found=parts_found, timestamp=ts,
                    ro_open_date="", **kwargs)

    def test_counts_match_classification(self, tmp_path):
        rows = [
            self._row_ts(24),              # PENDING
            self._row_ts(72),              # HARD_BLOCK (replace, no parts)
            self._row_ts(240),             # DIVERGE
            self._row_ts(72, desc="check oil", parts_found="False"),  # REVIEW
        ]
        csv_path = self._make_csv(rows, tmp_path)
        counts = run_analysis(csv_path, tmp_path / "out", KEYWORDS, NEGATORS, now=NOW)
        assert counts[WINDOW_PENDING]  == 1
        assert counts[WINDOW_HARD]     == 1
        assert counts[WINDOW_DIVERGE]  == 1
        assert counts[WINDOW_REVIEW]   == 1

    def test_output_files_created(self, tmp_path):
        rows = [self._row_ts(72)]
        csv_path = self._make_csv(rows, tmp_path)
        out_dir = tmp_path / "out"
        run_analysis(csv_path, out_dir, KEYWORDS, NEGATORS, now=NOW)
        assert (out_dir / "Closing_Blockers.csv").exists()
        assert (out_dir / "Human_Review_Queue.csv").exists()
        assert (out_dir / "Pattern_Learning_Queue.csv").exists()
        assert (out_dir / "Audit_Summary.txt").exists()

    def test_hard_block_row_written_to_closing_blockers(self, tmp_path):
        rows = [self._row_ts(72, ro_number="999111")]
        csv_path = self._make_csv(rows, tmp_path)
        out_dir = tmp_path / "out"
        run_analysis(csv_path, out_dir, KEYWORDS, NEGATORS, now=NOW)
        with open(out_dir / "Closing_Blockers.csv", newline="") as fh:
            content = fh.read()
        assert "999111" in content

    def test_missing_required_column_raises(self, tmp_path):
        p = tmp_path / "bad.csv"
        with open(p, "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=["Description", "Parts_Found"])
            w.writeheader()
            w.writerow({"Description": "test", "Parts_Found": "False"})
        with pytest.raises(ValueError, match="missing required columns"):
            run_analysis(p, tmp_path / "out", KEYWORDS, NEGATORS, now=NOW)

    def test_all_pending_no_blocking_rows(self, tmp_path):
        rows = [self._row_ts(h) for h in (1, 12, 24, 47)]
        csv_path = self._make_csv(rows, tmp_path)
        counts = run_analysis(csv_path, tmp_path / "out", KEYWORDS, NEGATORS, now=NOW)
        assert counts[WINDOW_PENDING] == 4
        assert counts[WINDOW_HARD]    == 0


# ---------------------------------------------------------------------------
# validate_headers
# ---------------------------------------------------------------------------

class TestValidateHeaders:
    def test_valid_headers_no_exception(self):
        validate_headers(list(
            {"Timestamp", "RO_Number", "Labor_ID",
             "Description", "Parts_Found", "Sequence", "RO_Open_Date"}
        ))

    def test_missing_column_raises_value_error(self):
        with pytest.raises(ValueError, match="missing required columns"):
            validate_headers(["Timestamp", "RO_Number", "Labor_ID", "Description"])
