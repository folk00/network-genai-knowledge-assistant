from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

SKIP_DIRS = {
    ".git",
    ".pytest_cache",
    "__pycache__",
    ".venv",
    "venv",
    "data/private",
    "real_data",
    "customer_data",
    "cert_question_banks",
}

RISKY_SUFFIXES = {
    ".docx",
    ".xlsx",
    ".pdf",
    ".pem",
    ".key",
    ".p12",
    ".pfx",
}

RISKY_NAME_PATTERNS = [
    re.compile(r"(?i)quiz_cache"),
    re.compile(r"(?i)real_?quiz"),
    re.compile(r"(?i)question_?bank"),
    re.compile(r"(?i)dump"),
    re.compile(r"(?i)customer"),
    re.compile(r"(?i)secret"),
    re.compile(r"(?i)credential"),
]

SECRET_TEXT_PATTERNS = [
    re.compile(r"\bsk-[A-Za-z0-9_-]{8,}\b"),
    re.compile(r"\bAKIA[0-9A-Z]{16}\b"),
    re.compile(r"(?i)-----BEGIN [A-Z ]*PRIVATE KEY-----"),
    re.compile(r"(?i)(password|api[_-]?key|secret|token)\s*[:=]\s*['\"]?[^'\"\r\n\t ]+"),
]


def main() -> int:
    findings: list[str] = []
    for path in ROOT.rglob("*"):
        if path.is_dir():
            continue
        rel = path.relative_to(ROOT).as_posix()
        if _is_skipped(rel):
            continue
        if path.suffix.lower() in RISKY_SUFFIXES:
            findings.append(f"risky extension: {rel}")
        if any(p.search(path.name) for p in RISKY_NAME_PATTERNS):
            findings.append(f"risky filename:  {rel}")
        if path.suffix.lower() in {".py", ".md", ".txt", ".csv", ".json", ".example"}:
            try:
                text = path.read_text(encoding="utf-8", errors="replace")
            except Exception:
                continue
            for pattern in SECRET_TEXT_PATTERNS:
                if _has_real_secret_match(pattern, text):
                    findings.append(f"possible secret: {rel}")
                    break

    if findings:
        print("Safety scan found files to review before GitHub:\n")
        for finding in findings:
            print(f"- {finding}")
        return 1

    print("Safety scan passed. No obvious private files or secrets found.")
    return 0


def _is_skipped(rel: str) -> bool:
    if rel == "scripts/safety_scan.py":
        return True
    parts = rel.split("/")
    for skip in SKIP_DIRS:
        skip_parts = skip.split("/")
        if parts[: len(skip_parts)] == skip_parts:
            return True
    return False


def _has_real_secret_match(pattern: re.Pattern[str], text: str) -> bool:
    for match in pattern.finditer(text):
        value = match.group(0)
        value_lower = value.lower()
        if "..." in value or "<" in value or value.endswith("="):
            continue
        if any(placeholder in value_lower for placeholder in (
            "dummy",
            "placeholder",
            "replace_with",
            "os.environ",
            "self._provider_artifact_token",
        )):
            continue
        if "api_key" in value_lower and ("\\" in value or "\n" in value):
            continue
        if value_lower.strip().endswith((":", ",")):
            continue
        # The scanner's own message strings are not secrets.
        if "possible secret" in value_lower:
            continue
        return True
    return False


if __name__ == "__main__":
    sys.exit(main())
