import unittest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from transient_artifact_classifier import classify_transient, is_transient_artifact


class TestTransientArtifactDefinition(unittest.TestCase):
    def test_log_extension_is_transient(self):
        self.assertTrue(is_transient_artifact("anything.log"))

    def test_pyc_extension_is_transient(self):
        self.assertTrue(is_transient_artifact("module.pyc"))

    def test_ini_is_not_transient(self):
        self.assertFalse(is_transient_artifact("settings.ini"))

    def test_txt_is_not_transient(self):
        self.assertFalse(is_transient_artifact("notes.txt"))

    def test_vbs_is_not_transient(self):
        self.assertFalse(is_transient_artifact("automation.vbs"))

    def test_transient_name_pattern_is_transient(self):
        result = classify_transient("debug_trace_report.txt")
        self.assertTrue(result["is_transient"])
        self.assertGreaterEqual(result["score"], 3)

    def test_allowlist_overrides_transient_detection(self):
        result = classify_transient(
            "reference_fixture.log",
            allowlist_patterns=[r"reference_fixture\.log$"],
        )
        self.assertFalse(result["is_transient"])
        self.assertEqual(result["reasons"], ["allowlisted"])


if __name__ == "__main__":
    unittest.main()
