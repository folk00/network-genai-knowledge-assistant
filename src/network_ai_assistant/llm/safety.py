from __future__ import annotations

import os
import re
from pathlib import Path


SECRET_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"\bsk-[A-Za-z0-9_-]{8,}\b"), "sk-***"),
    (re.compile(r"\bAKIA[0-9A-Z]{16}\b"), "AKIA***"),
    (re.compile(r"(?i)(api[_-]?key|token|secret|password)\s*[:=]\s*['\"]?[^'\"\r\n\t ]+"), r"\1=***"),
    (re.compile(r"(?is)-----BEGIN [A-Z ]*PRIVATE KEY-----.*?-----END [A-Z ]*PRIVATE KEY-----"), "[PRIVATE KEY REDACTED]"),
    (re.compile(r"(?i)(pre-shared-key|preshared[_ -]?key|psk)\s+[^ \n]+"), r"\1 ***"),
]


def redact_sensitive_text(text: str) -> str:
    """Best-effort redaction for prompts/logs.

    This is not a DLP product. It is a practical safety guard for demos.
    """

    redacted = str(text or "")
    for pattern, replacement in SECRET_PATTERNS:
        redacted = pattern.sub(replacement, redacted)
    return redacted


def load_env_file(path: Path | None = None) -> None:
    """Tiny .env loader using stdlib only.

    Existing environment variables win. This keeps the demo lightweight while
    avoiding a hard dependency on python-dotenv.
    """

    env_path = path or Path.cwd() / ".env"
    if not env_path.exists():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def live_calls_enabled() -> bool:
    return (os.environ.get("AI_LIVE") or "").strip().lower() in {"1", "true", "yes", "on"}
