import os
import re
from typing import Iterable, Optional


EXTENSION_RE = re.compile(
    r"\.(log|tmp|temp|out|err|pid|pid\.lock|coverage|pyc|pyo|class|o|obj)$",
    re.IGNORECASE,
)
TRANSIENT_NAME_RE = re.compile(
    r"(checkpoint|screendump|debug|trace|session|temp|tmp|backup|autosave)",
    re.IGNORECASE,
)


def _normalize_path(file_path: str) -> str:
    return file_path.replace("\\", "/")


def _allowlisted(file_path: str, allowlist_patterns: Optional[Iterable[str]]) -> bool:
    if not allowlist_patterns:
        return False
    normalized = _normalize_path(file_path)
    for pattern in allowlist_patterns:
        if pattern and re.search(pattern, normalized):
            return True
    return False


def classify_transient(
    file_path: str,
    allowlist_patterns: Optional[Iterable[str]] = None,
) -> dict:
    normalized = _normalize_path(file_path)
    score = 0
    reasons = []

    if _allowlisted(normalized, allowlist_patterns):
        return {"is_transient": False, "score": 0, "reasons": ["allowlisted"]}

    if EXTENSION_RE.search(normalized):
        score += 3
        reasons.append("extension")

    basename = os.path.basename(normalized)
    if TRANSIENT_NAME_RE.search(basename):
        score += 3
        reasons.append("name")

    return {"is_transient": score >= 3, "score": score, "reasons": reasons}


def is_transient_artifact(
    file_path: str,
    allowlist_patterns: Optional[Iterable[str]] = None,
) -> bool:
    return classify_transient(file_path, allowlist_patterns)["is_transient"]
