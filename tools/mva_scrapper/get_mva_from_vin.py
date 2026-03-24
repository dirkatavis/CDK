from __future__ import annotations

import configparser
import csv
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Optional

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


@dataclass(frozen=True)
class Settings:
    input_csv: Path
    output_csv: Path
    log_file: Path
    diagnostic_log_file: Path
    vin_column: str
    vin_prompt_text: str
    results_ready_text: str
    no_result_text: str
    mva_label_text: str
    mva_regex: str
    search_submit_key: str
    return_to_search_key: str
    command_wait_sec: int
    results_wait_sec: int
    poll_ms: int
    capture_screen_on_error: bool
    continue_on_error: bool


class BlueZoneClient:
    def __init__(self) -> None:
        if win32com is None:
            raise RuntimeError(
                "pywin32 is required to use BlueZone COM automation. Install with: pip install pywin32"
            )
        self._session = win32com.client.Dispatch("BZWhll.WhllObj")

    def connect(self) -> None:
        self._session.Connect("")

    def send_key(self, text: str) -> None:
        self._session.SendKey(text)

    def read_screen_all(self, length: int = 1920) -> str:
        if pythoncom is not None:
            try:
                buffer = win32com.client.VARIANT(
                    pythoncom.VT_BYREF | pythoncom.VT_BSTR,
                    " " * length,
                )
                self._session.ReadScreen(buffer, length, 1, 1)
                return str(buffer.value or "")
            except Exception:
                pass

        try:
            raw = self._session.ReadScreen(" " * length, length, 1, 1)
            if isinstance(raw, str):
                return raw
            if isinstance(raw, tuple):
                for item in raw:
                    if isinstance(item, str):
                        return item
        except Exception:
            return ""

        return ""


def get_repo_root() -> Path:
    base_path = os.environ.get("CDK_BASE", "").strip()
    if not base_path:
        raise RuntimeError("Invalid or missing CDK_BASE.")

    root = Path(base_path).resolve()
    if not root.is_dir():
        raise RuntimeError(f"Invalid CDK_BASE path: {root}")
    if not (root / ".cdkroot").exists():
        raise RuntimeError(f"Cannot find .cdkroot in base path: {root}")
    return root


def parse_bool(value: str, default: bool) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    return default


def resolve_path(repo_root: Path, value: str) -> Path:
    candidate = Path(value)
    if candidate.is_absolute():
        return candidate
    return (repo_root / candidate).resolve()


def load_settings(repo_root: Path) -> Settings:
    config_path = repo_root / "config" / "config.ini"
    parser = configparser.ConfigParser(interpolation=None)
    parser.read(config_path, encoding="utf-8")

    section = "GetMvaFromVin"
    if not parser.has_section(section):
        raise RuntimeError(f"Missing [{section}] section in {config_path}")

    def get_value(key: str, default: str = "") -> str:
        return parser.get(section, key, fallback=default).strip()

    return Settings(
        input_csv=resolve_path(repo_root, get_value("InputCSV")),
        output_csv=resolve_path(repo_root, get_value("OutputCSV")),
        log_file=resolve_path(repo_root, get_value("Log")),
        diagnostic_log_file=resolve_path(repo_root, get_value("DiagnosticLog")),
        vin_column=get_value("VinColumn", "VIN"),
        vin_prompt_text=get_value("VinPromptText", "VIN"),
        results_ready_text=get_value("ResultsReadyText", "MVA"),
        no_result_text=get_value("NoResultText", "NO VEHICLE FOUND"),
        mva_label_text=get_value("MvaLabelText", "MVA"),
        mva_regex=get_value("MvaRegex", r"MVA[: ]*([A-Z0-9\-]+)"),
        search_submit_key=get_value("SearchSubmitKey", "<NumpadEnter>"),
        return_to_search_key=get_value("ReturnToSearchKey", ""),
        command_wait_sec=int(get_value("CommandWaitSec", "10")),
        results_wait_sec=int(get_value("ResultsWaitSec", "15")),
        poll_ms=int(get_value("PollMs", "300")),
        capture_screen_on_error=parse_bool(get_value("CaptureScreenOnError", "true"), True),
        continue_on_error=parse_bool(get_value("ContinueOnError", "true"), True),
    )


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def configure_logger(log_file: Path) -> logging.Logger:
    ensure_parent(log_file)
    logger = logging.getLogger("get_mva_from_vin")
    logger.handlers.clear()
    logger.setLevel(logging.INFO)

    handler = logging.FileHandler(log_file, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    return logger


def write_diag(path: Path, title: str, screen_text: str) -> None:
    ensure_parent(path)
    clean_screen = screen_text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    with path.open("a", encoding="utf-8") as handle:
        handle.write(f"==== {time.strftime('%Y-%m-%d %H:%M:%S')} :: {title} ====\n")
        handle.write(clean_screen + "\n\n")


def wait_for_text(client: BlueZoneClient, text_to_find: str, timeout_sec: int, poll_ms: int) -> bool:
    deadline = time.monotonic() + timeout_sec
    needle = text_to_find.lower()

    while time.monotonic() < deadline:
        screen = client.read_screen_all()
        if needle in screen.lower():
            return True
        time.sleep(max(0, poll_ms) / 1000.0)

    return False


def wait_for_result_screen(client: BlueZoneClient, settings: Settings) -> Optional[str]:
    deadline = time.monotonic() + settings.results_wait_sec
    ready_tokens = [settings.results_ready_text.lower(), settings.mva_label_text.lower()]
    not_found_token = settings.no_result_text.lower()

    while time.monotonic() < deadline:
        screen = client.read_screen_all()
        lowered = screen.lower()

        if not_found_token in lowered:
            return screen
        if any(token and token in lowered for token in ready_tokens):
            return screen

        time.sleep(max(0, settings.poll_ms) / 1000.0)

    return None


def extract_mva(screen_text: str, mva_regex: str) -> str:
    match = re.search(mva_regex, screen_text, re.IGNORECASE)
    if not match:
        return ""
    if match.groups():
        return match.group(1).strip()
    return ""


def maybe_return_to_search(client: BlueZoneClient, settings: Settings) -> None:
    if not settings.return_to_search_key:
        return

    if wait_for_text(client, settings.vin_prompt_text, timeout_sec=1, poll_ms=settings.poll_ms):
        return

    client.send_key(settings.return_to_search_key)
    time.sleep(max(0, settings.poll_ms) / 1000.0)


def process_vin(
    client: BlueZoneClient,
    settings: Settings,
    logger: logging.Logger,
    vin_value: str,
) -> Dict[str, str]:
    logger.info("Processing VIN %s", vin_value)

    if not wait_for_text(client, settings.vin_prompt_text, settings.command_wait_sec, settings.poll_ms):
        return {
            "VIN": vin_value,
            "MVA": "",
            "Status": "ERROR",
            "Error": "VIN prompt not found before input",
        }

    client.send_key(f"{vin_value}{settings.search_submit_key}")
    time.sleep(max(0, settings.poll_ms) / 1000.0)

    result_screen = wait_for_result_screen(client, settings)
    if result_screen is None:
        return {
            "VIN": vin_value,
            "MVA": "",
            "Status": "TIMEOUT",
            "Error": "Timed out waiting for Compass result",
        }

    if settings.no_result_text.lower() in result_screen.lower():
        maybe_return_to_search(client, settings)
        return {
            "VIN": vin_value,
            "MVA": "",
            "Status": "NOT_FOUND",
            "Error": "No vehicle result returned",
        }

    mva_value = extract_mva(result_screen, settings.mva_regex)
    if not mva_value:
        maybe_return_to_search(client, settings)
        return {
            "VIN": vin_value,
            "MVA": "",
            "Status": "ERROR",
            "Error": "MVA not found on results screen",
        }

    logger.info("VIN %s -> MVA %s", vin_value, mva_value)
    maybe_return_to_search(client, settings)
    return {
        "VIN": vin_value,
        "MVA": mva_value,
        "Status": "OK",
        "Error": "",
    }


def iter_rows(input_csv: Path) -> Iterable[Dict[str, str]]:
    with input_csv.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            yield {k: (v or "") for k, v in row.items()}


def run() -> int:
    repo_root = get_repo_root()
    settings = load_settings(repo_root)
    logger = configure_logger(settings.log_file)

    logger.info("Starting get_mva_from_vin")
    logger.info("InputCSV=%s", settings.input_csv)
    logger.info("OutputCSV=%s", settings.output_csv)

    if not settings.input_csv.exists():
        logger.error("Input CSV not found: %s", settings.input_csv)
        return 1

    try:
        client = BlueZoneClient()
        client.connect()
    except Exception as exc:
        logger.error("Failed to connect to BlueZone/Compass session: %s", exc)
        return 1

    ensure_parent(settings.output_csv)

    total_rows = 0
    success_rows = 0
    failed_rows = 0

    with settings.output_csv.open("w", encoding="utf-8", newline="") as out_handle:
        fieldnames = ["VIN", "MVA", "Status", "Error"]
        writer = csv.DictWriter(out_handle, fieldnames=fieldnames)
        writer.writeheader()

        for row in iter_rows(settings.input_csv):
            total_rows += 1
            vin_value = row.get(settings.vin_column, "").strip()
            if not vin_value:
                failed_rows += 1
                writer.writerow({"VIN": "", "MVA": "", "Status": "SKIPPED", "Error": "Empty VIN"})
                continue

            result = process_vin(client, settings, logger, vin_value)
            writer.writerow(result)

            if result["Status"] == "OK":
                success_rows += 1
            else:
                failed_rows += 1
                logger.error("VIN %s failed: %s", vin_value, result["Error"])
                if settings.capture_screen_on_error:
                    write_diag(
                        settings.diagnostic_log_file,
                        f"VIN {vin_value} failure screen",
                        client.read_screen_all(),
                    )
                if not settings.continue_on_error:
                    logger.error("Stopping run due to failure and ContinueOnError=false")
                    break

    logger.info(
        "Completed get_mva_from_vin. Total=%s, Success=%s, Failed=%s",
        total_rows,
        success_rows,
        failed_rows,
    )
    return 0


if __name__ == "__main__":
    sys.exit(run())
