"""
Tests for apps/labor_parts_analyzer/discover_keywords.py
Run from repo root:  python -m pytest apps/labor_parts_analyzer/tests/test_discover_keywords.py -v
"""

import sys
import csv
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from discover_keywords import (
    tokenize,
    compute_token_rates,
    classify_token,
    build_report,
    write_report,
    KEYWORD_THRESHOLD,
    NEGATOR_THRESHOLD,
    LABEL_KEYWORD,
    LABEL_NEGATOR,
    LABEL_AMBIGUOUS,
)


# ---------------------------------------------------------------------------
# tokenize
# ---------------------------------------------------------------------------

class TestTokenize:
    def test_basic_split(self):
        assert tokenize("replace oil filter") == ["replace", "oil", "filter"]

    def test_drops_single_chars(self):
        assert "a" not in tokenize("a replace b filter")

    def test_drops_pure_numbers(self):
        assert "123" not in tokenize("replace 123 filter")

    def test_lowercases(self):
        assert "REPLACE" not in tokenize("REPLACE OIL")
        assert "replace" in tokenize("REPLACE OIL")

    def test_strips_punctuation(self):
        tokens = tokenize("replace, oil. filter!")
        assert tokens == ["replace", "oil", "filter"]

    def test_empty_string(self):
        assert tokenize("") == []

    def test_two_char_tokens_kept(self):
        assert "do" in tokenize("do the work")


# ---------------------------------------------------------------------------
# compute_token_rates
# ---------------------------------------------------------------------------

class TestComputeTokenRates:
    def _rows(self, entries):
        return [{"Description": d, "Parts_Found": pf} for d, pf in entries]

    def test_single_keyword_row_rate_one(self):
        rows = self._rows([("replace filter", "True")])
        rates = compute_token_rates(rows)
        assert rates["replace"]["parts_true_rate"] == pytest.approx(1.0)

    def test_single_negator_row_rate_zero(self):
        rows = self._rows([("inspect tires", "False")])
        rates = compute_token_rates(rows)
        assert rates["inspect"]["parts_true_rate"] == pytest.approx(0.0)

    def test_mixed_rate(self):
        rows = self._rows([
            ("replace filter", "True"),
            ("replace tires",  "False"),
        ])
        rates = compute_token_rates(rows)
        assert rates["replace"]["parts_true_rate"] == pytest.approx(0.5)
        assert rates["replace"]["total"] == 2

    def test_total_count_correct(self):
        rows = self._rows([("replace filter", "True")] * 3)
        rates = compute_token_rates(rows)
        assert rates["replace"]["total"] == 3

    def test_parts_true_count_correct(self):
        rows = self._rows([
            ("replace", "True"),
            ("replace", "True"),
            ("replace", "False"),
        ])
        rates = compute_token_rates(rows)
        assert rates["replace"]["parts_true"] == 2


# ---------------------------------------------------------------------------
# classify_token
# ---------------------------------------------------------------------------

class TestClassifyToken:
    def test_at_keyword_threshold(self):
        assert classify_token(KEYWORD_THRESHOLD) == LABEL_KEYWORD

    def test_above_keyword_threshold(self):
        assert classify_token(1.0) == LABEL_KEYWORD

    def test_at_negator_threshold(self):
        assert classify_token(NEGATOR_THRESHOLD) == LABEL_NEGATOR

    def test_below_negator_threshold(self):
        assert classify_token(0.0) == LABEL_NEGATOR

    def test_middle_is_ambiguous(self):
        assert classify_token(0.5) == LABEL_AMBIGUOUS

    def test_just_below_keyword_is_ambiguous(self):
        assert classify_token(KEYWORD_THRESHOLD - 0.01) == LABEL_AMBIGUOUS

    def test_just_above_negator_is_ambiguous(self):
        assert classify_token(NEGATOR_THRESHOLD + 0.01) == LABEL_AMBIGUOUS


# ---------------------------------------------------------------------------
# build_report
# ---------------------------------------------------------------------------

class TestBuildReport:
    def _rates(self, entries):
        """entries: [(token, total, parts_true)]"""
        return {
            tok: {
                "total":           total,
                "parts_true":      pt,
                "parts_true_rate": pt / total if total > 0 else 0.0,
            }
            for tok, total, pt in entries
        }

    def test_min_count_filters_low_freq(self):
        rates = self._rates([("rare", 2, 2), ("common", 10, 10)])
        report = build_report(rates, min_count=5)
        tokens = [r["Token"] for r in report]
        assert "rare" not in tokens
        assert "common" in tokens

    def test_sorted_by_rate_descending(self):
        rates = self._rates([
            ("low",  10, 1),    # 0.10
            ("high", 10, 9),    # 0.90
            ("mid",  10, 5),    # 0.50
        ])
        report = build_report(rates, min_count=1)
        tokens = [r["Token"] for r in report]
        assert tokens == ["high", "mid", "low"]

    def test_label_assigned_correctly(self):
        rates = self._rates([
            ("kw",  10, 10),   # 1.0  → KEYWORD
            ("neg", 10, 1),    # 0.10 → NEGATOR
            ("amb", 10, 5),    # 0.50 → AMBIGUOUS
        ])
        report = build_report(rates, min_count=1)
        labels = {r["Token"]: r["Label"] for r in report}
        assert labels["kw"]  == LABEL_KEYWORD
        assert labels["neg"] == LABEL_NEGATOR
        assert labels["amb"] == LABEL_AMBIGUOUS

    def test_empty_rates_returns_empty_report(self):
        assert build_report({}, min_count=1) == []


# ---------------------------------------------------------------------------
# write_report
# ---------------------------------------------------------------------------

class TestWriteReport:
    def _sample_report(self):
        return [
            {"Token": "replace", "Total": 10, "Parts_True": 9,
             "Parts_True_Rate": 0.9, "Label": LABEL_KEYWORD},
            {"Token": "inspect", "Total": 8,  "Parts_True": 1,
             "Parts_True_Rate": 0.125, "Label": LABEL_NEGATOR},
            {"Token": "drain",   "Total": 6,  "Parts_True": 3,
             "Parts_True_Rate": 0.5,  "Label": LABEL_AMBIGUOUS},
        ]

    def test_discovery_csv_created(self, tmp_path):
        write_report(self._sample_report(), tmp_path)
        assert (tmp_path / "Keyword_Discovery.csv").exists()

    def test_keyword_csv_contains_only_keywords(self, tmp_path):
        write_report(self._sample_report(), tmp_path)
        with open(tmp_path / "Keywords.csv", newline="") as fh:
            rows = list(csv.DictReader(fh))
        assert all(r["Label"] == LABEL_KEYWORD for r in rows)
        assert len(rows) == 1

    def test_negator_csv_contains_only_negators(self, tmp_path):
        write_report(self._sample_report(), tmp_path)
        with open(tmp_path / "Negators.csv", newline="") as fh:
            rows = list(csv.DictReader(fh))
        assert all(r["Label"] == LABEL_NEGATOR for r in rows)

    def test_keywords_list_txt_contains_tokens(self, tmp_path):
        write_report(self._sample_report(), tmp_path)
        content = (tmp_path / "Keywords_List.txt").read_text()
        assert "replace" in content

    def test_negators_list_txt_contains_tokens(self, tmp_path):
        write_report(self._sample_report(), tmp_path)
        content = (tmp_path / "Negators_List.txt").read_text()
        assert "inspect" in content

    def test_ambiguous_csv_created(self, tmp_path):
        write_report(self._sample_report(), tmp_path)
        assert (tmp_path / "Ambiguous.csv").exists()
