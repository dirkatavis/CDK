"""
Keyword Discovery — discover_keywords.py
Part of the CDK DMS Automation Suite — apps/labor_parts_analyzer/discover_keywords.py

Reads Master_Labor_Log.csv and computes per-token frequency rates split by
Parts_Found=True vs Parts_Found=False.  Tokens that appear predominantly in
Parts_Found=True rows are promoted to KEYWORD; those that appear predominantly
in Parts_Found=False rows become NEGATOR candidates.

Classification thresholds:
  KEYWORD   — parts_true_rate >= 0.70   (token predicts parts present)
  NEGATOR   — parts_true_rate <= 0.15   (token predicts parts absent)
  AMBIGUOUS — everything else

Usage:
    python discover_keywords.py [--config path/to/config.ini] [--min-count N]
"""

import argparse
import configparser
import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
KEYWORD_THRESHOLD = 0.70
NEGATOR_THRESHOLD = 0.15
DEFAULT_MIN_COUNT = 5

LABEL_KEYWORD   = "KEYWORD"
LABEL_NEGATOR   = "NEGATOR"
LABEL_AMBIGUOUS = "AMBIGUOUS"

REQUIRED_COLUMNS = {"Description", "Parts_Found"}


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def tokenize(text: str) -> list[str]:
    """
    Lowercase, strip punctuation, and split on whitespace.
    Single-character tokens and pure numbers are dropped.
    """
    lower = text.lower()
    # Replace common punctuation with spaces
    cleaned = re.sub(r"[^a-z0-9\s]", " ", lower)
    tokens = cleaned.split()
    return [t for t in tokens if len(t) > 1 and not t.isdigit()]


def compute_token_rates(rows: list[dict]) -> dict[str, dict]:
    """
    Returns a dict keyed by token:
      {token: {"total": int, "parts_true": int, "parts_true_rate": float}}
    """
    total_counts: Counter = Counter()
    true_counts:  Counter = Counter()

    for row in rows:
        desc   = row.get("Description", "")
        is_true = row.get("Parts_Found", "False").strip().lower() == "true"

        for tok in tokenize(desc):
            total_counts[tok] += 1
            if is_true:
                true_counts[tok] += 1

    result = {}
    for tok, total in total_counts.items():
        parts_true = true_counts.get(tok, 0)
        result[tok] = {
            "total":           total,
            "parts_true":      parts_true,
            "parts_true_rate": parts_true / total if total > 0 else 0.0,
        }
    return result


def classify_token(parts_true_rate: float) -> str:
    if parts_true_rate >= KEYWORD_THRESHOLD:
        return LABEL_KEYWORD
    if parts_true_rate <= NEGATOR_THRESHOLD:
        return LABEL_NEGATOR
    return LABEL_AMBIGUOUS


def build_report(token_rates: dict, min_count: int = DEFAULT_MIN_COUNT) -> list[dict]:
    """
    Build flat list of report rows, filtered to tokens with >= min_count occurrences.
    Sorted by parts_true_rate descending.
    """
    rows = []
    for tok, stats in token_rates.items():
        if stats["total"] < min_count:
            continue
        label = classify_token(stats["parts_true_rate"])
        rows.append({
            "Token":           tok,
            "Total":           stats["total"],
            "Parts_True":      stats["parts_true"],
            "Parts_True_Rate": round(stats["parts_true_rate"], 4),
            "Label":           label,
        })
    rows.sort(key=lambda r: r["Parts_True_Rate"], reverse=True)
    return rows


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------

def write_report(report_rows: list[dict], output_dir: Path) -> None:
    """Write Keyword_Discovery.csv (full) and per-label CSV/txt splits."""
    output_dir.mkdir(parents=True, exist_ok=True)

    all_cols = ["Token", "Total", "Parts_True", "Parts_True_Rate", "Label"]

    # Full discovery report
    _write_csv(output_dir / "Keyword_Discovery.csv", all_cols, report_rows)

    # Per-label splits
    keywords   = [r for r in report_rows if r["Label"] == LABEL_KEYWORD]
    negators   = [r for r in report_rows if r["Label"] == LABEL_NEGATOR]
    ambiguous  = [r for r in report_rows if r["Label"] == LABEL_AMBIGUOUS]

    _write_csv(output_dir / "Keywords.csv",   all_cols, keywords)
    _write_csv(output_dir / "Negators.csv",   all_cols, negators)
    _write_csv(output_dir / "Ambiguous.csv",  all_cols, ambiguous)

    # Plain word lists
    _write_word_list(output_dir / "Keywords_List.txt",
                     [r["Token"] for r in keywords])
    _write_word_list(output_dir / "Negators_List.txt",
                     [r["Token"] for r in negators])


def _write_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _write_word_list(path: Path, words: list[str]) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        for word in words:
            fh.write(word + "\n")


# ---------------------------------------------------------------------------
# Config / CLI
# ---------------------------------------------------------------------------

def _find_config(provided: str | None) -> Path:
    if provided:
        return Path(provided)
    here = Path(__file__).resolve().parent
    for candidate in (here, here.parent, here.parent.parent):
        p = candidate / "config" / "config.ini"
        if p.exists():
            return p
    raise FileNotFoundError(
        "Cannot locate config/config.ini. "
        "Pass --config <path> or run from the repo root."
    )


def load_rows(input_csv: Path) -> list[dict]:
    with open(input_csv, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        fieldnames = list(reader.fieldnames or [])
        missing = REQUIRED_COLUMNS - set(f.strip() for f in fieldnames)
        if missing:
            raise ValueError(
                f"Input CSV is missing required columns: {sorted(missing)}"
            )
        return list(reader)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Discover keywords from labor log")
    parser.add_argument("--config",    default=None, help="Path to config.ini")
    parser.add_argument("--min-count", type=int, default=DEFAULT_MIN_COUNT,
                        help=f"Minimum token occurrences (default: {DEFAULT_MIN_COUNT})")
    args = parser.parse_args(argv)

    config_path = _find_config(args.config)
    cfg         = configparser.ConfigParser()
    cfg.read(str(config_path))

    repo_root = config_path.parent.parent

    input_csv  = repo_root / cfg.get("LaborPartsAnalyzer", "InputCSV")
    output_dir = repo_root / cfg.get("LaborPartsAnalyzer", "OutputDir",
                                      fallback="runtime/data/Keyword_Analysis")

    if not input_csv.exists():
        print(f"ERROR: Input CSV not found: {input_csv}", file=sys.stderr)
        sys.exit(1)

    try:
        rows = load_rows(input_csv)
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    rates  = compute_token_rates(rows)
    report = build_report(rates, min_count=args.min_count)
    write_report(report, output_dir)

    keywords  = sum(1 for r in report if r["Label"] == LABEL_KEYWORD)
    negators  = sum(1 for r in report if r["Label"] == LABEL_NEGATOR)
    ambiguous = sum(1 for r in report if r["Label"] == LABEL_AMBIGUOUS)

    print(f"Discovery complete.  {len(report)} tokens analyzed (min_count={args.min_count}).")
    print(f"  {LABEL_KEYWORD}:   {keywords}")
    print(f"  {LABEL_NEGATOR}:  {negators}")
    print(f"  {LABEL_AMBIGUOUS}: {ambiguous}")
    print(f"Outputs written to: {output_dir}")


if __name__ == "__main__":
    main()
