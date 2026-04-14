"""
Labor Parts Analyzer — analyze.py
Part of the CDK DMS Automation Suite — apps/labor_parts_analyzer/analyze.py

Reads Master_Labor_Log.csv (produced by tools/labor_parts_scraper.vbs) and
classifies each L-line row into one of three temporal windows:

  PENDING_LEAD_TIME    — RO age   0–48 h   (parts may still be on order)
  HARD_BLOCK           — RO age  48 h–7 d AND description matches a keyword
                         AND Parts_Found is False → likely a logic gap
  DIVERGENCE_LEARNING  — RO age   8+ days  (patterns worth learning from)

Rows that do not fall into a named bucket are written to the Human_Review_Queue.

Usage:
    python analyze.py [--config path/to/config.ini]
"""

import argparse
import configparser
import csv
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
REQUIRED_COLUMNS = {
    "Timestamp", "RO_Number", "Labor_ID",
    "Description", "Parts_Found", "Sequence", "RO_Open_Date",
}

WINDOW_PENDING  = "PENDING_LEAD_TIME"
WINDOW_HARD     = "HARD_BLOCK"
WINDOW_DIVERGE  = "DIVERGENCE_LEARNING"
WINDOW_REVIEW   = "HUMAN_REVIEW"

PENDING_HOURS   = 48
HARD_MAX_DAYS   = 7
DIVERGE_MIN_DAYS = 8


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_config(config_path: Path) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg.read(str(config_path))
    return cfg


def parse_keyword_list(raw: str) -> list[str]:
    """Split comma-separated keyword string into lowercased tokens."""
    return [k.strip().lower() for k in raw.split(",") if k.strip()]


def classify_description(desc: str, keywords: list[str], negators: list[str]) -> bool:
    """
    Return True if the description contains a keyword AND does NOT contain a negator.
    Substring matching is used (not word-boundary) to keep this simple and fast.
    """
    lower = desc.lower()
    has_keyword = any(kw in lower for kw in keywords)
    has_negator  = any(neg in lower for neg in negators)
    return has_keyword and not has_negator


def parse_ro_open_date(ro_open_date: str, now: datetime) -> datetime | None:
    """
    Parse CDK DDMMMYY format (e.g. '05NOV25') into a datetime.
    Returns None if the string cannot be parsed.
    """
    if not ro_open_date or ro_open_date.upper() in ("UNKNOWN", ""):
        return None
    try:
        # CDK stores 2-digit year (e.g. 25 → 2025).  strptime %y handles this.
        return datetime.strptime(ro_open_date.strip().upper(), "%d%b%y")
    except ValueError:
        return None


def parse_timestamp(ts_str: str) -> datetime | None:
    """Parse the Timestamp column written by Now() in VBScript."""
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.strptime(ts_str.strip(), fmt)
        except ValueError:
            pass
    return None


def ro_age_hours(row: dict, now: datetime) -> float | None:
    """
    Compute RO age in hours.
    Priority: RO_Open_Date (parsed from CDK DDMMMYY) → fallback to Timestamp.
    Returns None if neither can be parsed.
    """
    dt = parse_ro_open_date(row.get("RO_Open_Date", ""), now)
    if dt is None:
        dt = parse_timestamp(row.get("Timestamp", ""))
    if dt is None:
        return None
    return (now - dt).total_seconds() / 3600.0


def classify_row(row: dict, keywords: list[str], negators: list[str], now: datetime) -> str:
    """Assign the row to one of the four buckets."""
    age_h = ro_age_hours(row, now)
    if age_h is None:
        return WINDOW_REVIEW

    desc        = row.get("Description", "")
    parts_found = row.get("Parts_Found", "False").strip().lower() == "true"
    matches_kw  = classify_description(desc, keywords, negators)

    if age_h <= PENDING_HOURS:
        return WINDOW_PENDING

    age_d = age_h / 24.0
    if age_d <= HARD_MAX_DAYS:
        if matches_kw and not parts_found:
            return WINDOW_HARD
        return WINDOW_REVIEW

    # age >= 8 days
    return WINDOW_DIVERGE


# ---------------------------------------------------------------------------
# I/O helpers
# ---------------------------------------------------------------------------

def validate_headers(fieldnames: list[str]) -> None:
    missing = REQUIRED_COLUMNS - set(f.strip() for f in fieldnames)
    if missing:
        raise ValueError(
            f"Input CSV is missing required columns: {sorted(missing)}\n"
            f"Expected: {sorted(REQUIRED_COLUMNS)}\n"
            f"Found:    {sorted(fieldnames)}"
        )


def make_output_writers(output_dir: Path):
    """
    Return a dict of {bucket_name: (file_obj, csv.DictWriter)} for each output.
    Caller must close the file objects.
    """
    writers = {}
    files   = {}

    output_dir.mkdir(parents=True, exist_ok=True)

    file_map = {
        WINDOW_HARD:    "Closing_Blockers.csv",
        WINDOW_REVIEW:  "Human_Review_Queue.csv",
        WINDOW_DIVERGE: "Pattern_Learning_Queue.csv",
    }
    row_cols = list(REQUIRED_COLUMNS) + ["Window", "RO_Age_Hours"]

    for bucket, filename in file_map.items():
        fh = open(output_dir / filename, "w", newline="", encoding="utf-8")
        writer = csv.DictWriter(fh, fieldnames=row_cols, extrasaction="ignore")
        writer.writeheader()
        files[bucket]   = fh
        writers[bucket] = writer

    return files, writers


def run_analysis(input_csv: Path, output_dir: Path,
                 keywords: list[str], negators: list[str],
                 now: datetime | None = None) -> dict:
    """
    Main entry point.  Returns summary dict with row counts per bucket.
    Raises ValueError on header mismatch.
    """
    if now is None:
        now = datetime.now()

    counts = {WINDOW_PENDING: 0, WINDOW_HARD: 0, WINDOW_DIVERGE: 0, WINDOW_REVIEW: 0}

    with open(input_csv, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        validate_headers(list(reader.fieldnames or []))

        files, writers = make_output_writers(output_dir)

        try:
            for row in reader:
                bucket  = classify_row(row, keywords, negators, now)
                counts[bucket] += 1

                if bucket != WINDOW_PENDING:
                    age_h = ro_age_hours(row, now)
                    out_row = dict(row)
                    out_row["Window"]       = bucket
                    out_row["RO_Age_Hours"] = f"{age_h:.1f}" if age_h is not None else ""
                    writers[bucket].writerow(out_row)
        finally:
            for fh in files.values():
                fh.close()

    # Write audit summary
    _write_audit_summary(output_dir, counts, input_csv, now)

    return counts


def _write_audit_summary(output_dir: Path, counts: dict,
                          input_csv: Path, now: datetime) -> None:
    total = sum(counts.values())
    lines = [
        f"Labor Parts Analysis — Audit Summary",
        f"Run at:       {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"Input file:   {input_csv}",
        f"Total rows:   {total}",
        f"",
        f"  {WINDOW_PENDING:<30} {counts[WINDOW_PENDING]:>6}",
        f"  {WINDOW_HARD:<30} {counts[WINDOW_HARD]:>6}  ← Closing_Blockers.csv",
        f"  {WINDOW_DIVERGE:<30} {counts[WINDOW_DIVERGE]:>6}  ← Pattern_Learning_Queue.csv",
        f"  {WINDOW_REVIEW:<30} {counts[WINDOW_REVIEW]:>6}  ← Human_Review_Queue.csv",
    ]
    (output_dir / "Audit_Summary.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def _find_config(provided: str | None) -> Path:
    """
    Locate config.ini.  If --config not given, walk upward from this script's
    directory looking for config/config.ini (handles running from repo root or
    from the apps/labor_parts_analyzer subdirectory).
    """
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


def main(argv=None):
    parser = argparse.ArgumentParser(description="Labor Parts Analyzer")
    parser.add_argument("--config", default=None, help="Path to config.ini")
    args = parser.parse_args(argv)

    config_path = _find_config(args.config)
    cfg         = load_config(config_path)

    repo_root = config_path.parent.parent

    input_csv  = repo_root / cfg.get("LaborPartsAnalyzer", "InputCSV")
    output_dir = repo_root / cfg.get("LaborPartsAnalyzer", "OutputDir",
                                      fallback="runtime/data/analysis")

    keywords_raw = cfg.get("LaborPartsAnalyzer", "High_Confidence_Keywords",
                            fallback="replace,install,new,kit,assembly")
    negators_raw = cfg.get("LaborPartsAnalyzer", "Negators",
                            fallback="inspect,check,verify,test,adjust,measure,clean,diagnose,drain")

    keywords = parse_keyword_list(keywords_raw)
    negators = parse_keyword_list(negators_raw)

    if not input_csv.exists():
        print(f"ERROR: Input CSV not found: {input_csv}", file=sys.stderr)
        sys.exit(1)

    try:
        counts = run_analysis(input_csv, output_dir, keywords, negators)
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    total = sum(counts.values())
    print(f"Analysis complete.  {total} rows processed.")
    print(f"  {WINDOW_HARD}:    {counts[WINDOW_HARD]}")
    print(f"  {WINDOW_DIVERGE}: {counts[WINDOW_DIVERGE]}")
    print(f"  {WINDOW_PENDING}: {counts[WINDOW_PENDING]}")
    print(f"  {WINDOW_REVIEW}:  {counts[WINDOW_REVIEW]}")
    print(f"Outputs written to: {output_dir}")


if __name__ == "__main__":
    main()
