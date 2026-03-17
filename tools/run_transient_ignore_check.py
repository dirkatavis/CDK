import argparse
import configparser
from pathlib import Path
from datetime import datetime

from transient_artifact_classifier import classify_transient


def _read_multiline_list(config: configparser.ConfigParser, section: str, option: str) -> list[str]:
    if not config.has_option(section, option):
        return []
    raw = config.get(section, option)
    return [line.strip() for line in raw.splitlines() if line.strip()]


def _read_expected_actions(config: configparser.ConfigParser) -> dict[str, str]:
    entries = _read_multiline_list(config, "expectations", "actions")
    parsed: dict[str, str] = {}
    for entry in entries:
        if "|" not in entry:
            continue
        file_name, action = entry.split("|", 1)
        parsed[file_name.strip()] = action.strip()
    return parsed


def _recommended_action(expected_transient: bool, classified_transient: bool) -> str:
    if expected_transient and classified_transient:
        return "ignore-in-repo-keep-local"
    if expected_transient and not classified_transient:
        return "review-before-ignore"
    if not expected_transient and classified_transient:
        return "add-allowlist-pattern"
    return "track-file"


def _action_hint(file_name: str, action: str) -> str:
    if action == "ignore-in-repo-keep-local":
        return f"hint=add .gitignore rule; if already tracked use: git rm --cached \"{file_name}\""
    if action == "add-allowlist-pattern":
        return "hint=add regex to .transient-allowlist"
    if action == "review-before-ignore":
        return "hint=manual review: classifier uncertain for expected transient file"
    return "hint=keep tracked"


def main() -> int:
    parser = argparse.ArgumentParser(description="Run transient-artifact checks from an INI file.")
    parser.add_argument(
        "--ini",
        default="tools/transient_ignore_samples.ini",
        help="Path to INI file containing sample filenames.",
    )
    parser.add_argument(
        "--out",
        default="",
        help="Optional file path to append the report output.",
    )
    parser.add_argument(
        "--validate-expectations",
        action="store_true",
        help="Validate against [expectations] actions in the INI (default is blind mode).",
    )
    args = parser.parse_args()

    ini_path = Path(args.ini)
    if not ini_path.exists():
        print(f"INI file not found: {ini_path}")
        return 2

    config = configparser.ConfigParser()
    config.read(ini_path, encoding="utf-8")

    transient_files = _read_multiline_list(config, "samples", "transient")
    tracked_files = _read_multiline_list(config, "samples", "tracked")
    allowlist_patterns = _read_multiline_list(config, "allowlist", "patterns")
    expected_actions = _read_expected_actions(config) if args.validate_expectations else {}

    output_lines: list[str] = []
    output_lines.append(f"Run: {datetime.now().isoformat(timespec='seconds')}")
    output_lines.append(f"Using INI: {ini_path}")

    output_lines.append("")
    output_lines.append("Transient candidates:")
    transient_failures = 0
    action_failures = 0
    for file_name in transient_files:
        result = classify_transient(file_name, allowlist_patterns=allowlist_patterns)
        status = "PASS" if result["is_transient"] else "FAIL"
        if status == "FAIL":
            transient_failures += 1
        action = _recommended_action(expected_transient=True, classified_transient=result["is_transient"])
        if args.validate_expectations:
            expected_action = expected_actions.get(file_name, "")
            action_status = "PASS" if (not expected_action or expected_action == action) else "FAIL"
            if action_status == "FAIL":
                action_failures += 1
            output_lines.append(
                f"  [{status}/{action_status}] {file_name} -> score={result['score']} reasons={','.join(result['reasons'])} action={action} expected={expected_action or '-'} {_action_hint(file_name, action)}"
            )
        else:
            output_lines.append(
                f"  {file_name} -> transient={result['is_transient']} score={result['score']} reasons={','.join(result['reasons'])} action={action} {_action_hint(file_name, action)}"
            )

    output_lines.append("")
    output_lines.append("Tracked candidates:")
    tracked_failures = 0
    for file_name in tracked_files:
        result = classify_transient(file_name, allowlist_patterns=allowlist_patterns)
        status = "PASS" if not result["is_transient"] else "FAIL"
        if status == "FAIL":
            tracked_failures += 1
        action = _recommended_action(expected_transient=False, classified_transient=result["is_transient"])
        if args.validate_expectations:
            expected_action = expected_actions.get(file_name, "")
            action_status = "PASS" if (not expected_action or expected_action == action) else "FAIL"
            if action_status == "FAIL":
                action_failures += 1
            output_lines.append(
                f"  [{status}/{action_status}] {file_name} -> score={result['score']} reasons={','.join(result['reasons'])} action={action} expected={expected_action or '-'} {_action_hint(file_name, action)}"
            )
        else:
            output_lines.append(
                f"  {file_name} -> transient={result['is_transient']} score={result['score']} reasons={','.join(result['reasons'])} action={action} {_action_hint(file_name, action)}"
            )

    output_lines.append("")
    if args.validate_expectations:
        total_failures = transient_failures + tracked_failures + action_failures
        output_lines.append(
            f"Summary: {total_failures} failure(s) [classification={transient_failures + tracked_failures}, action={action_failures}]"
        )
    else:
        output_lines.append("Summary: blind mode (no expectation validation)")
    output_lines.append("Safety note: ignore decisions mean exclude from Git while keeping local files.")

    report_text = "\n".join(output_lines)
    print(report_text)

    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with out_path.open("a", encoding="utf-8") as handle:
            handle.write(report_text)
            handle.write("\n\n")

    if args.validate_expectations:
        return 1 if (transient_failures + tracked_failures + action_failures) else 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
