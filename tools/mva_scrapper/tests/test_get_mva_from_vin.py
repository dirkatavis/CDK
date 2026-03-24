"""Unit and integration tests for get_mva_from_vin module.

Tests cover:
  1. Pure utility functions (extract_mva, parse_bool, resolve_path, iter_rows)
  2. Mocked BlueZone integration (process_vin, wait_for_text, wait_for_result_screen)
  3. Configuration & path validation (load_settings, get_repo_root)
"""

from __future__ import annotations

import configparser
import csv
import logging
import os
import tempfile
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, Mock, patch

import pytest

# Import the module under test
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from get_mva_from_vin import (
    BlueZoneClient,
    Settings,
    extract_mva,
    get_repo_root,
    iter_rows,
    load_settings,
    parse_bool,
    process_vin,
    resolve_path,
    wait_for_result_screen,
    wait_for_text,
)


# ============================================================================
# PART 1: Pure Utility Function Tests
# ============================================================================


class TestExtractMva:
    """Tests for the extract_mva regex extraction function."""

    def test_extract_mva_basic(self) -> None:
        """Extract MVA from standard format."""
        screen = "MVA: ABC123-XYZ"
        result = extract_mva(screen, r"MVA[: ]*([A-Z0-9\-]+)")
        assert result == "ABC123-XYZ"

    def test_extract_mva_no_space(self) -> None:
        """Extract MVA when colon and space are absent."""
        screen = "MVAABC123"
        result = extract_mva(screen, r"MVA[: ]*([A-Z0-9\-]+)")
        assert result == "ABC123"

    def test_extract_mva_case_insensitive(self) -> None:
        """Extract MVA from mixed-case screen text."""
        screen = "result mva: FORD001"
        result = extract_mva(screen, r"MVA[: ]*([A-Z0-9\-]+)")
        assert result == "FORD001"

    def test_extract_mva_with_hyphens(self) -> None:
        """Extract MVA containing hyphens."""
        screen = "MVA: ABC-123-XYZ"
        result = extract_mva(screen, r"MVA[: ]*([A-Z0-9\-]+)")
        assert result == "ABC-123-XYZ"

    def test_extract_mva_not_found(self) -> None:
        """Return empty string when MVA not found."""
        screen = "No matching pattern here"
        result = extract_mva(screen, r"MVA[: ]*([A-Z0-9\-]+)")
        assert result == ""

    def test_extract_mva_no_capture_group(self) -> None:
        """Return empty string when regex has no capture group."""
        screen = "MVA: ABC123"
        result = extract_mva(screen, r"MVA")
        assert result == ""

    def test_extract_mva_custom_regex(self) -> None:
        """Extract MVA with custom regex pattern."""
        screen = "Vehicle Code: VHL12345"
        result = extract_mva(screen, r"Code: ([A-Z0-9]+)")
        assert result == "VHL12345"


class TestParseBool:
    """Tests for the parse_bool string-to-boolean converter."""

    @pytest.mark.parametrize(
        "value,expected",
        [
            ("true", True),
            ("True", True),
            ("TRUE", True),
            ("1", True),
            ("yes", True),
            ("y", True),
            ("on", True),
            ("false", False),
            ("False", False),
            ("FALSE", False),
            ("0", False),
            ("no", False),
            ("n", False),
            ("off", False),
        ],
    )
    def test_parse_bool_valid_inputs(self, value: str, expected: bool) -> None:
        """Parse valid bool strings correctly."""
        assert parse_bool(value, default=None) == expected

    def test_parse_bool_with_whitespace(self) -> None:
        """Trim whitespace before parsing."""
        assert parse_bool("  true  ", default=False) is True
        assert parse_bool("  false  ", default=True) is False

    def test_parse_bool_invalid_uses_default_true(self) -> None:
        """Use default=True for unrecognized input."""
        assert parse_bool("maybe", default=True) is True

    def test_parse_bool_invalid_uses_default_false(self) -> None:
        """Use default=False for unrecognized input."""
        assert parse_bool("unknown", default=False) is False

    def test_parse_bool_empty_string_uses_default(self) -> None:
        """Empty string uses provided default."""
        assert parse_bool("", default=True) is True
        assert parse_bool("", default=False) is False


class TestResolvePath:
    """Tests for the resolve_path function."""

    def test_resolve_path_absolute(self) -> None:
        """Keep absolute paths unchanged."""
        abs_path = Path("C:/abs/path/file.txt")
        repo_root = Path("C:/repo")
        result = resolve_path(repo_root, str(abs_path))
        assert result == abs_path.resolve()

    def test_resolve_path_relative(self) -> None:
        """Resolve relative paths from repo root."""
        repo_root = Path("C:/repo")
        relative = "config/config.ini"
        result = resolve_path(repo_root, relative)
        assert result == (repo_root / relative).resolve()

    def test_resolve_path_with_dots(self) -> None:
        """Handle relative paths with .. correctly."""
        repo_root = Path("C:/repo/subdir")
        relative = "../config/file.ini"
        result = resolve_path(repo_root, relative)
        assert result == (repo_root / relative).resolve()

    def test_resolve_path_windows_backslash(self) -> None:
        """Handle Windows-style backslash paths."""
        repo_root = Path("C:\\repo")
        relative = "config\\config.ini"
        result = resolve_path(repo_root, relative)
        assert str(result).endswith("config.ini")


class TestIterRows:
    """Tests for CSV row iteration with encoding handling."""

    def test_iter_rows_simple(self) -> None:
        """Iterate basic CSV rows."""
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline=""
        ) as f:
            writer = csv.DictWriter(f, fieldnames=["VIN", "Status"])
            writer.writeheader()
            writer.writerow({"VIN": "ABC123", "Status": "OK"})
            writer.writerow({"VIN": "XYZ789", "Status": "ERROR"})
            temp_path = Path(f.name)

        try:
            rows = list(iter_rows(temp_path))
            assert len(rows) == 2
            assert rows[0]["VIN"] == "ABC123"
            assert rows[0]["Status"] == "OK"
            assert rows[1]["VIN"] == "XYZ789"
            assert rows[1]["Status"] == "ERROR"
        finally:
            temp_path.unlink()

    def test_iter_rows_empty_fields(self) -> None:
        """Handle empty fields correctly."""
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline=""
        ) as f:
            writer = csv.DictWriter(f, fieldnames=["VIN", "MVA"])
            writer.writeheader()
            writer.writerow({"VIN": "ABC123", "MVA": ""})
            writer.writerow({"VIN": "XYZ789", "MVA": None})
            temp_path = Path(f.name)

        try:
            rows = list(iter_rows(temp_path))
            assert rows[0]["MVA"] == ""
            assert rows[1]["MVA"] == ""
        finally:
            temp_path.unlink()

    def test_iter_rows_with_bom(self) -> None:
        """Handle UTF-8 BOM in CSV file."""
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8-sig", suffix=".csv", delete=False, newline=""
        ) as f:
            writer = csv.DictWriter(f, fieldnames=["VIN", "Status"])
            writer.writeheader()
            writer.writerow({"VIN": "ABC123", "Status": "OK"})
            temp_path = Path(f.name)

        try:
            rows = list(iter_rows(temp_path))
            assert len(rows) == 1
            assert rows[0]["VIN"] == "ABC123"
        finally:
            temp_path.unlink()

    def test_iter_rows_nonexistent_file(self) -> None:
        """Raise error for missing CSV file."""
        nonexistent = Path("/nonexistent/file.csv")
        with pytest.raises(FileNotFoundError):
            list(iter_rows(nonexistent))


# ============================================================================
# PART 2: Mocked BlueZone Integration Tests
# ============================================================================


class TestWaitForText:
    """Tests for the wait_for_text blocking function."""

    def test_wait_for_text_found_immediately(self) -> None:
        """Return True when text is found on first check."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "VIN prompt found here"

        result = wait_for_text(mock_client, "VIN prompt", timeout_sec=5, poll_ms=100)

        assert result is True
        mock_client.read_screen_all.assert_called()

    def test_wait_for_text_case_insensitive(self) -> None:
        """Search is case-insensitive."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "SEARCH RESULTS READY"

        result = wait_for_text(mock_client, "search results", timeout_sec=5, poll_ms=100)

        assert result is True

    def test_wait_for_text_timeout(self) -> None:
        """Return False when timeout expires."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "No match here"

        result = wait_for_text(mock_client, "expected text", timeout_sec=0.1, poll_ms=50)

        assert result is False

    def test_wait_for_text_found_after_retries(self) -> None:
        """Return True when text appears after multiple polls."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.side_effect = [
            "first screen",
            "second screen",
            "finally found the text",
        ]

        result = wait_for_text(mock_client, "found the text", timeout_sec=2, poll_ms=100)

        assert result is True


class TestWaitForResultScreen:
    """Tests for the wait_for_result_screen function."""

    def test_wait_for_result_screen_found(self) -> None:
        """Return screen when MVA result is found."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "MVA: ABC123-XYZ"

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=5,
            results_wait_sec=5,
            poll_ms=100,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        result = wait_for_result_screen(mock_client, settings)

        assert result is not None
        assert "ABC123" in result

    def test_wait_for_result_screen_not_found(self) -> None:
        """Return screen with no-result message."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "NO VEHICLE FOUND"

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=5,
            results_wait_sec=5,
            poll_ms=100,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        result = wait_for_result_screen(mock_client, settings)

        assert result is not None
        assert "NO VEHICLE" in result

    def test_wait_for_result_screen_timeout(self) -> None:
        """Return None on timeout."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "Loading..."

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=5,
            results_wait_sec=0.1,
            poll_ms=50,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        result = wait_for_result_screen(mock_client, settings)

        assert result is None


class TestProcessVin:
    """Tests for the main process_vin orchestration function."""

    def test_process_vin_success(self) -> None:
        """Process VIN successfully and extract MVA."""
        mock_client = Mock(spec=BlueZoneClient)

        # First call: check for VIN prompt (must succeed)
        # This is mocked to always succeed with a 2-second timeout
        def screen_side_effect() -> str:
            return "Enter VIN: "

        mock_client.read_screen_all.side_effect = [
            "Enter VIN: ",  # prompt found
            "MVA: FORD001",  # result screen
        ]

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="Enter VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=5,
            results_wait_sec=5,
            poll_ms=100,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        logger = logging.getLogger("test")

        result = process_vin(mock_client, settings, logger, "ABC123")

        assert result["VIN"] == "ABC123"
        assert result["MVA"] == "FORD001"
        assert result["Status"] == "OK"
        assert result["Error"] == ""

    def test_process_vin_prompt_not_found(self) -> None:
        """Handle case when VIN prompt is never found."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.return_value = "Loading screen..."

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="Enter VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=0.1,
            results_wait_sec=5,
            poll_ms=50,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        logger = logging.getLogger("test")

        result = process_vin(mock_client, settings, logger, "ABC123")

        assert result["VIN"] == "ABC123"
        assert result["MVA"] == ""
        assert result["Status"] == "ERROR"
        assert "VIN prompt" in result["Error"]

    def test_process_vin_no_vehicle_found(self) -> None:
        """Handle 'no vehicle found' response."""
        mock_client = Mock(spec=BlueZoneClient)
        mock_client.read_screen_all.side_effect = [
            "Enter VIN: ",  # prompt found
            "NO VEHICLE FOUND",  # result: no match
        ]

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="Enter VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=5,
            results_wait_sec=5,
            poll_ms=100,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        logger = logging.getLogger("test")

        result = process_vin(mock_client, settings, logger, "BADVIN00")

        assert result["VIN"] == "BADVIN00"
        assert result["MVA"] == ""
        assert result["Status"] == "NOT_FOUND"

    def test_process_vin_timeout_waiting_for_result(self) -> None:
        """Handle timeout while waiting for result screen."""
        mock_client = Mock(spec=BlueZoneClient)
        # Return different values: first time prompt found, then always stuck
        mock_client.read_screen_all.side_effect = [
            "Enter VIN: ",  # prompt found (first call)
        ] + ["Loading... (never finishes)"] * 10  # stuck screen (retries during polling)

        settings = Settings(
            input_csv=Path("/tmp/in.csv"),
            output_csv=Path("/tmp/out.csv"),
            log_file=Path("/tmp/log.txt"),
            diagnostic_log_file=Path("/tmp/diag.txt"),
            vin_column="VIN",
            vin_prompt_text="Enter VIN",
            results_ready_text="MVA",
            no_result_text="NO VEHICLE",
            mva_label_text="MVA",
            mva_regex=r"MVA[: ]*([A-Z0-9\-]+)",
            search_submit_key="<Return>",
            return_to_search_key="",
            command_wait_sec=5,
            results_wait_sec=0.1,
            poll_ms=50,
            capture_screen_on_error=False,
            continue_on_error=True,
        )

        logger = logging.getLogger("test")

        result = process_vin(mock_client, settings, logger, "ABC123")

        assert result["VIN"] == "ABC123"
        assert result["MVA"] == ""
        assert result["Status"] == "TIMEOUT"


# ============================================================================
# PART 3: Configuration & Path Validation Tests
# ============================================================================


class TestGetRepoRoot:
    """Tests for the get_repo_root function."""

    def test_get_repo_root_valid(self) -> None:
        """Return repo root when CDK_BASE and .cdkroot exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            (tmp_path / ".cdkroot").touch()

            with patch.dict(os.environ, {"CDK_BASE": str(tmp_path)}):
                result = get_repo_root()
                assert result == tmp_path

    def test_get_repo_root_missing_env(self) -> None:
        """Raise error when CDK_BASE is not set."""
        with patch.dict(os.environ, {}, clear=True):
            with pytest.raises(RuntimeError, match="Invalid or missing CDK_BASE"):
                get_repo_root()

    def test_get_repo_root_missing_cdkroot_marker(self) -> None:
        """Raise error when .cdkroot marker is missing."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            # Don't create .cdkroot

            with patch.dict(os.environ, {"CDK_BASE": str(tmp_path)}):
                with pytest.raises(RuntimeError, match=".cdkroot"):
                    get_repo_root()

    def test_get_repo_root_invalid_path(self) -> None:
        """Raise error when CDK_BASE points to nonexistent directory."""
        with patch.dict(os.environ, {"CDK_BASE": "/nonexistent/path"}):
            with pytest.raises(RuntimeError, match="Invalid CDK_BASE path"):
                get_repo_root()


class TestLoadSettings:
    """Tests for the load_settings configuration loader."""

    def test_load_settings_valid(self) -> None:
        """Load valid configuration successfully."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            (tmp_path / ".cdkroot").touch()
            config_dir = tmp_path / "config"
            config_dir.mkdir()

            config_file = config_dir / "config.ini"
            config_file.write_text(
                """[GetMvaFromVin]
InputCSV=data/input.csv
OutputCSV=data/output.csv
Log=logs/mva.log
DiagnosticLog=logs/mva_diag.log
VinColumn=VIN
VinPromptText=Enter VIN
ResultsReadyText=MVA
NoResultText=NO VEHICLE
MvaLabelText=MVA
MvaRegex=MVA[: ]*([A-Z0-9\\-]+)
SearchSubmitKey=<Return>
ReturnToSearchKey=
CommandWaitSec=10
ResultsWaitSec=15
PollMs=300
CaptureScreenOnError=true
ContinueOnError=true
"""
            )

            settings = load_settings(tmp_path)

            assert settings.vin_column == "VIN"
            assert settings.command_wait_sec == 10
            assert settings.results_wait_sec == 15
            assert settings.poll_ms == 300
            assert settings.capture_screen_on_error is True
            assert settings.continue_on_error is True

    def test_load_settings_missing_section(self) -> None:
        """Raise error when GetMvaFromVin section is missing."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            config_dir = tmp_path / "config"
            config_dir.mkdir()

            config_file = config_dir / "config.ini"
            config_file.write_text("[OtherSection]\nKey=Value\n")

            with pytest.raises(RuntimeError, match="GetMvaFromVin"):
                load_settings(tmp_path)

    def test_load_settings_defaults(self) -> None:
        """Use default values for optional configuration keys."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            config_dir = tmp_path / "config"
            config_dir.mkdir()

            config_file = config_dir / "config.ini"
            config_file.write_text(
                """[GetMvaFromVin]
InputCSV=data/input.csv
OutputCSV=data/output.csv
Log=logs/mva.log
DiagnosticLog=logs/mva_diag.log
"""
            )

            settings = load_settings(tmp_path)

            # Verify defaults
            assert settings.vin_column == "VIN"
            assert settings.results_ready_text == "MVA"
            assert settings.no_result_text == "NO VEHICLE FOUND"
            assert settings.command_wait_sec == 10
            assert settings.results_wait_sec == 15

    def test_load_settings_parse_bool_flags(self) -> None:
        """Parse true/false configuration flags correctly."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            config_dir = tmp_path / "config"
            config_dir.mkdir()

            config_file = config_dir / "config.ini"
            config_file.write_text(
                """[GetMvaFromVin]
InputCSV=data/input.csv
OutputCSV=data/output.csv
Log=logs/mva.log
DiagnosticLog=logs/mva_diag.log
CaptureScreenOnError=false
ContinueOnError=yes
"""
            )

            settings = load_settings(tmp_path)

            assert settings.capture_screen_on_error is False
            assert settings.continue_on_error is True
