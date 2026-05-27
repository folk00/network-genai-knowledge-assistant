#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ANS-C01 Quiz GUI (PySide6) - v2

Fixes requested:
- Options are selectable (radio buttons for single, checkboxes for multi)
- Practice mode: NO timer; instant feedback + running score
- After wrong answers: store into a "Wrong Bank" and allow practicing only those
- Shuffle options so you don't memorize position
- Shuffle questions optional
- Show explanations (from DOCX explanation or overrides notes/expl)

Dependencies:
  pip install PySide6 python-docx qdarkstyle

Run:
  python ans_c01_quiz_gui.py

Data files (optional):
- Overrides CSV columns supported:
    id, doc_answer, corrected, notes, my_explanation
  (extra columns are ignored)
"""

from __future__ import annotations

import csv
import os
import base64
import hashlib
import json
import random
import re
import sys
import traceback
import faulthandler
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date as _date, timedelta as _timedelta
import html
from dataclasses import dataclass, field
from pathlib import Path

try:
    _REPO_ROOT = Path(__file__).resolve().parents[2]
    _SRC_ROOT = _REPO_ROOT / "src"
    if _SRC_ROOT.exists() and str(_SRC_ROOT) not in sys.path:
        sys.path.insert(0, str(_SRC_ROOT))
except Exception:
    _REPO_ROOT = Path.cwd()

# ---- UI theme ----
# Set to True if you want the dark theme (qdarkstyle). Default is system/light.
USE_DARK_THEME = True

# ---- crash diagnostics (helps when Qt hard-crashes) ----
# Writes to ans_c01_quiz_crash.log next to this script (or CWD fallback).
try:
    _CRASH_LOG_PATH = Path(__file__).with_name("ans_c01_quiz_crash.log")
except Exception:
    _CRASH_LOG_PATH = Path.cwd() / "ans_c01_quiz_crash.log"

try:
    _crash_fh = open(_CRASH_LOG_PATH, "a", buffering=1, encoding="utf-8", errors="replace")
    _crash_fh.write("\n\n===== START %s =====\n" % datetime.now().isoformat(timespec="seconds"))
    faulthandler.enable(_crash_fh)
except Exception:
    _crash_fh = None

def _install_excepthook() -> None:
    def _hook(exc_type, exc, tb):
        try:
            msg = "".join(traceback.format_exception(exc_type, exc, tb))
            with open(_CRASH_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
                f.write("\n[PYTHON EXCEPTION %s]\n" % datetime.now().isoformat(timespec="seconds"))
                f.write(msg)
        except Exception:
            pass
        # keep default behavior (prints to stderr if available)
        sys.__excepthook__(exc_type, exc, tb)
    sys.excepthook = _hook

_install_excepthook()
# ------------------------------------------------------
from typing import Dict, List, Optional, Set, FrozenSet, Tuple

from PySide6.QtCore import Qt, QTimer, QThread, QObject, Signal, Slot, QSettings, QUrl
from PySide6.QtGui import QFont, QIcon, QCursor, QDesktopServices, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QProgressDialog,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTextEdit,
    QTextBrowser,
    QToolButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

try:
    from docx import Document  # python-docx
except Exception:
    Document = None


_DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
NEW_QID_START = 306
NEW_QID_END = 322

AI_PROVIDER_CLAUDE = "claude"
AI_PROVIDER_OPENAI = "openai"
AI_PROVIDER_LABELS = {
    AI_PROVIDER_CLAUDE: "Claude",
    AI_PROVIDER_OPENAI: "OpenAI",
}
AI_PROVIDER_BY_LABEL = {v: k for k, v in AI_PROVIDER_LABELS.items()}
AI_MODELS_BY_PROVIDER = {
    AI_PROVIDER_CLAUDE: [
        "(Claude default)",
        "claude-opus-4-7",
        "claude-sonnet-4-6",
    ],
    AI_PROVIDER_OPENAI: [
        "gpt-5.4-mini",
        "gpt-5.4",
        "gpt-5.5",
        "gpt-5.2-pro",
        "gpt-4.1",
        "gpt-4o",
    ],
}
AI_REPORT_FILES = {
    "coach": ("ai_coach_reports.md", "ai_coach_reports_openai.md"),
    "deep": ("deep_review_reports.md", "deep_review_reports_openai.md"),
    "nuclear": ("nuclear_reports.md", "nuclear_reports_openai.md"),
    "meta": ("ai_meta_coach.md", "ai_meta_coach_openai.md"),
    "meta_hash": ("ai_meta_coach.md.sha256", "ai_meta_coach_openai.md.sha256"),
    "meta_history": ("ai_meta_history.json", "ai_meta_history_openai.json"),
    "diagram_index": ("diagram_reports.md", "diagram_reports_openai.md"),
    "teach_index": ("teach_zero_reports.md", "teach_zero_reports_openai.md"),
}


def _paragraph_text(node: ET.Element) -> str:
    return "".join((t.text or "") for t in node.findall(".//w:t", _DOCX_NS)).strip()


def _cell_text(node: ET.Element) -> str:
    parts: List[str] = []
    for child in node:
        local = child.tag.rsplit("}", 1)[-1]
        if local == "p":
            text = _paragraph_text(child)
            if text:
                parts.append(text)
        elif local == "tbl":
            for line in _table_lines(child):
                if line:
                    parts.append(line)
    return " ".join(parts).strip()


def _table_lines(node: ET.Element) -> List[str]:
    lines: List[str] = []
    for row in node.findall("./w:tr", _DOCX_NS):
        cells: List[str] = []
        for cell in row.findall("./w:tc", _DOCX_NS):
            text = _cell_text(cell)
            if text:
                cells.append(text)
        if not cells:
            continue

        first = cells[0]
        if len(cells) >= 2 and re.match(r"^[A-F]\s*[\.\)\:]\s*$", first, re.IGNORECASE):
            lines.append(f"{first} {cells[1]}".strip())
            for extra in cells[2:]:
                extra = extra.strip()
                if extra:
                    lines.append(extra)
            continue

        merged = " ".join(cells).strip()
        if merged:
            lines.append(merged)
    return lines


def _read_docx_lines(path: Path) -> List[str]:
    try:
        with zipfile.ZipFile(path, "r") as zf:
            with zf.open("word/document.xml") as f:
                xml_bytes = f.read()
    except Exception as exc:
        raise RuntimeError(
            "Could not read DOCX contents. Install python-docx or verify the file is a valid .docx"
        ) from exc

    try:
        root = ET.fromstring(xml_bytes)
    except Exception as exc:
        raise RuntimeError("DOCX XML could not be parsed") from exc

    body = root.find("w:body", _DOCX_NS)
    if body is None:
        return []

    lines: List[str] = []
    for child in body:
        local = child.tag.rsplit("}", 1)[-1]
        if local == "p":
            line = _paragraph_text(child)
            if line:
                lines.append(line)
        elif local == "tbl":
            lines.extend(_table_lines(child))
    return lines


# -----------------------------
# Parsing
# -----------------------------
_WORD_TO_NUM = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6}
_CHOOSE_RE = re.compile(r"\bChoose\s+(one|two|three|four|five|six|\d+)\b", re.IGNORECASE)

_Q_RE = re.compile(
    r"^(?:QUESTION|Q)\s*(?:NO\.?|NUMBER)?\s*[:#\-]?\s*(\d+)\b",
    re.IGNORECASE,
)
# Allow: A. / A) / A:  (some dumps vary)
_OPT_RE = re.compile(r"^([A-F])\s*[\.\)\:]\s*(.*)$")
_ANS_LINE_RE = re.compile(
    r"^(Answer|Correct Answer)\s*:\s*([A-F]{1,6}(?:[,\s]+[A-F]{1,6})*)\b",
    re.IGNORECASE,
)
_ANS_ANY_RE = re.compile(
    r"\b(?:Answer|Correct Answer)\s*:\s*([A-F]{1,6}(?:[,\s]+[A-F]{1,6})*)\b",
    re.IGNORECASE,
)
_EXPL_RE = re.compile(r"^Explanation\s*:?", re.IGNORECASE)

_JUNK_PATS = [
    re.compile(r"passleader\.com", re.IGNORECASE),
    re.compile(r"Get Latest", re.IGNORECASE),
    re.compile(r"^Page\s+\d+", re.IGNORECASE),
    re.compile(r"^\d+$"),
]


def _is_junk(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return True
    for pat in _JUNK_PATS:
        if pat.search(s):
            return True
    return False


def _normalize_answer(ans: str) -> str:
    letters = re.findall(r"[A-F]", (ans or "").upper())
    out: List[str] = []
    for ch in letters:
        if ch not in out:
            out.append(ch)
    return "".join(out)


def _split_combined_paragraphs(lines: List[str]) -> List[str]:
    out: List[str] = []
    for line in lines:
        s = (line or "").strip()
        if not s:
            continue
        # OCR/text-extraction glitch in Q182 option B: "ALB. Upload" was fused.
        s = s.replace("ALUpload the certificate", "ALB. Upload the certificate")
        while True:
            m = re.search(r"\bQUESTION\s+\d+\b", s, re.IGNORECASE)
            if m and m.start() != 0:
                out.append(s[: m.start()].strip())
                s = s[m.start() :].strip()
                continue
            break
        out.append(s)
    return [x for x in out if x and x.strip()]


@dataclass
class Question:
    qid: int
    stem: str
    options: Dict[str, str]  # insertion-ordered
    answer_doc: str
    explanation: str = ""
    pattern_id: str = ""
    tags: List[str] = field(default_factory=list)

    def expected_count(self) -> Optional[int]:
        m = _CHOOSE_RE.search(self.stem or "")
        if not m:
            return None
        tok = m.group(1).lower()
        if tok in _WORD_TO_NUM:
            return _WORD_TO_NUM[tok]
        try:
            return int(tok)
        except Exception:
            return None

    def effective_answer(self, overrides: Dict[int, str]) -> str:
        a = overrides.get(self.qid, "").strip().upper()
        answer = _normalize_answer(a) if a else self.answer_doc
        # Validate answer letters are in available options
        valid_letters = set(self.options.keys())
        filtered = "".join([ch for ch in answer if ch in valid_letters])
        return filtered if filtered else answer


def parse_docx(path: Path) -> List[Question]:
    raw_lines = _read_docx_lines(path)

    if not raw_lines:
        raise RuntimeError("No readable text found in DOCX")

    lines = _split_combined_paragraphs(raw_lines)

    qs: List[Question] = []
    cur_id: Optional[int] = None
    stem_parts: List[str] = []
    expl_parts: List[str] = []
    options: Dict[str, str] = {}
    answer_doc: str = ""
    last_opt: Optional[str] = None
    in_expl: bool = False

    def flush():
        nonlocal cur_id, stem_parts, options, answer_doc, expl_parts
        if cur_id is None:
            return
        qs.append(
            Question(
                qid=cur_id,
                stem=" ".join(stem_parts).strip(),
                options=dict(options),
                answer_doc=answer_doc.strip(),
                explanation=" ".join(expl_parts).strip(),
            )
        )

    for raw in lines:
        line = raw.strip()
        if _is_junk(line):
            continue

        m = _Q_RE.match(line)
        if m:
            flush()
            cur_id = int(m.group(1))
            stem_parts = []
            expl_parts = []
            options = {}
            answer_doc = ""
            last_opt = None
            in_expl = False
            continue

        if cur_id is None:
            continue

        m = _ANS_LINE_RE.match(line)
        if m:
            answer_doc = _normalize_answer(m.group(2))
            in_expl = False
            after = line[m.end() :].strip()
            if after:
                if _EXPL_RE.match(after):
                    in_expl = True
                    after2 = _EXPL_RE.sub("", after).strip()
                    if after2:
                        expl_parts.append(after2)
                else:
                    expl_parts.append(after)
            continue

        m = _EXPL_RE.match(line)
        if m:
            in_expl = True
            rest = _EXPL_RE.sub("", line).strip()
            if rest:
                expl_parts.append(rest)
            continue

        m = _OPT_RE.match(line)
        if m:
            letter = m.group(1).upper()
            txt = (m.group(2) or "").strip()

            m2 = _ANS_ANY_RE.search(txt)
            if m2 and not answer_doc:
                answer_doc = _normalize_answer(m2.group(1))
                txt = txt[: m2.start()].rstrip()

            options[letter] = txt
            last_opt = letter
            in_expl = False
            continue

        # embedded answer in other lines
        m2 = _ANS_ANY_RE.search(line)
        if m2 and not answer_doc:
            answer_doc = _normalize_answer(m2.group(1))
            line = line[: m2.start()].rstrip()
            if not line:
                continue

        if in_expl:
            expl_parts.append(line)
        elif options and not answer_doc and last_opt:
            # continuation of option text
            options[last_opt] = (options[last_opt] + " " + line).strip()
        elif options and answer_doc:
            # some docs continue explanation without label
            expl_parts.append(line)
        else:
            stem_parts.append(line)

    flush()
    qs.sort(key=lambda x: x.qid)
    return qs


def _compute_file_sha256(path: Path) -> str:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ""


def _compute_questions_signature(questions: List[Question]) -> str:
    try:
        h = hashlib.sha256()
        for q in questions or []:
            h.update(f"QID:{int(q.qid)}\n".encode("utf-8", "replace"))
            h.update((q.stem or "").encode("utf-8", "replace"))
            h.update(b"\n")
            for letter, text in (q.options or {}).items():
                h.update(f"{letter}:{text}\n".encode("utf-8", "replace"))
            h.update(f"ANSWER:{q.answer_doc or ''}\n".encode("utf-8", "replace"))
            h.update(f"PATTERN:{q.pattern_id or ''}\n".encode("utf-8", "replace"))
            for tag in (q.tags or []):
                h.update(f"TAG:{tag}\n".encode("utf-8", "replace"))
            h.update(b"---\n")
        return h.hexdigest()
    except Exception:
        return ""


def load_overrides_csv(path: Path) -> Tuple[
    Dict[int, str], Dict[int, str], Dict[int, str], Dict[int, str], Dict[int, List[str]]
]:
    """Load a CSV with optional per-question fixes + metadata.

    Supported columns (case-insensitive, flexible names):
      - id / qid / question
      - corrected / corrected_answer / answer
      - notes
      - my_explanation / explanation
      - pattern_id / pattern
      - tags  (comma or semicolon separated)

    Returns:
      overrides[qid] -> corrected answer letters (e.g. 'C' or 'A,D')
      notes[qid] -> notes
      my_expl[qid] -> your explanation (preferred over DOCX explanation)
      patterns[qid] -> pattern_id
      tags_map[qid] -> list of tags
    """
    overrides: Dict[int, str] = {}
    notes: Dict[int, str] = {}
    my_expl: Dict[int, str] = {}
    patterns: Dict[int, str] = {}
    tags_map: Dict[int, List[str]] = {}

    if not path.exists():
        return overrides, notes, my_expl, patterns, tags_map

    def norm_tags(v: str) -> List[str]:
        raw = (v or "").strip()
        if not raw:
            return []
        # allow commas or semicolons
        parts = re.split(r"[;,]", raw)
        out: List[str] = []
        seen = set()
        for p in parts:
            t = p.strip()
            if not t:
                continue
            key = t.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(t)
        return out

    with path.open("r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not row:
                continue

            # id
            qid_raw = (
                row.get("id")
                or row.get("question")
                or row.get("qid")
                or row.get("ID")
                or row.get("Question")
                or row.get("QID")
            )
            if not qid_raw:
                continue
            try:
                qid = int(str(qid_raw).strip())
            except Exception:
                continue

            # corrected answer
            ans = row.get("corrected") or row.get("answer") or row.get("corrected_answer") or ""
            ans_n = _normalize_answer(str(ans).strip())
            if ans_n:
                overrides[qid] = ans_n

            # notes + explanation
            n = (row.get("notes") or "").strip()
            if n:
                notes[qid] = n

            me = (row.get("my_explanation") or row.get("explanation") or "").strip()
            if me:
                my_expl[qid] = me

            # pattern + tags
            pid = (row.get("pattern_id") or row.get("pattern") or "").strip()
            if pid:
                patterns[qid] = pid

            tg = norm_tags(str(row.get("tags") or ""))
            if tg:
                tags_map[qid] = tg

    return overrides, notes, my_expl, patterns, tags_map


def export_overrides_template(path: Path, qs: List[Question]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["id", "doc_answer", "corrected", "notes", "my_explanation", "pattern_id", "tags"])
        for q in qs:
            w.writerow([q.qid, q.answer_doc, q.answer_doc, "", "", q.pattern_id, ";".join(q.tags)])




# -----------------------------
# Pattern inference (optional)
# -----------------------------
def infer_pattern_and_tags(q: Question) -> Tuple[str, List[str]]:
    """Best-effort auto label. You can override via Overrides CSV.

    Returns (pattern_id, tags). Empty values mean 'unknown'.
    """
    s = (q.stem or "").lower()

    def has(*words: str) -> bool:
        return any(w.lower() in s for w in words)

    # PrivateLink / Endpoints
    if has("overlapping", "overlap") and has("privatelink", "endpoint service", "interface endpoint", "cross-account", "cross account"):
        return "PRIVATELINK_OVERLAP", ["PrivateLink", "OverlappingCIDR"]
    if has("private connectivity", "cross-account", "cross account") and has("endpoint service", "privatelink", "interface endpoint"):
        return "PRIVATELINK_CROSS_ACCOUNT", ["PrivateLink", "InterfaceEndpoint"]
    if has("s3") and has("without nat", "no nat", "without internet", "private", "no internet") and has("endpoint"):
        # Often gateway endpoint, but question may still be testing interface endpoints; label as endpoints family
        return "ENDPOINT_S3_PRIVATE", ["Endpoints", "S3"]

    # NAT
    if has("errorportallocation", "port allocation", "55,000", "55000"):
        return "NAT_PORT_EXHAUSTION", ["NAT", "Egress"]
    if has("nat gateway") and has("timeout", "idle") and has("keepalive", "keep alive", "tcp keepalive"):
        return "NAT_IDLE_TIMEOUT", ["NAT", "TCP"]

    # DNS / Route 53 / Resolver
    if has("route 53 resolver", "resolver") and has("inbound", "outbound"):
        return "R53_RESOLVER_INBOUND_OUTBOUND", ["DNS", "Route53Resolver"]
    if has("dns firewall") or has("resolver dns firewall"):
        return "R53_RESOLVER_DNS_FIREWALL", ["DNS", "Route53Resolver"]
    if has("evaluate target health") and has("alias") and has("alb"):
        return "R53_ALIAS_EVALUATE_HEALTH_ALB", ["DNS", "Route53", "ALB"]

    # Global routing
    if has("global accelerator"):
        return "GLOBAL_ACCELERATOR", ["Global", "GlobalAccelerator"]
    if has("cloudfront") and has("failover", "origin failover"):
        return "CLOUDFRONT_FAILOVER", ["CloudFront"]

    # DX / TGW
    if has("direct connect") and has("transit gateway", "tgw", "dxgw", "direct connect gateway", "transit vif", "site link", "sitelink"):
        return "DX_TGW", ["DirectConnect", "TGW"]

    # Security / Inspection
    if has("network firewall", "aws network firewall"):
        return "NETWORK_FIREWALL", ["Security", "NetworkFirewall"]
    if has("gateway load balancer", "gwlb"):
        return "GWLB", ["Security", "GWLB"]

    return "", []


def apply_patterns_from_overrides(questions: List[Question], patterns: Dict[int, str], tags_map: Dict[int, List[str]]) -> None:
    for q in questions:
        if patterns.get(q.qid):
            q.pattern_id = patterns[q.qid]
        if tags_map.get(q.qid):
            q.tags = tags_map[q.qid]
# -----------------------------
# Worker
# -----------------------------
class LoadWorker(QObject):
    finished = Signal(object, object)  # (questions, error)
    progress = Signal(int)

    def __init__(self, docx_path: Path):
        super().__init__()
        self.docx_path = docx_path

    @Slot()
    def run(self) -> None:
        try:
            self.progress.emit(10)
            qs = parse_docx(self.docx_path)
            self.progress.emit(95)
            self.finished.emit(qs, None)
        except Exception as e:
            self.finished.emit(None, str(e))


class AICoachWorker(QObject):
    """Runs quiz_ai_coach.analyze_wrong_answers in a QThread."""
    finished = Signal(str, str)        # (markdown_report, error_or_empty)
    progress_text = Signal(str)        # streaming status for progress dialog

    def __init__(self, items: List[Dict], model: Optional[str] = None, provider: str = "claude", timeout_sec: float = 180.0):
        super().__init__()
        self.items = items
        self.model = model
        self.provider = provider
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.worker")
        try:
            from quiz_ai_coach import analyze_wrong_answers, _LOG_PATH
            log.info("worker starting, items=%d, log=%s", len(self.items), _LOG_PATH)
            md = analyze_wrong_answers(
                self.items,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
            )
            log.info("worker got %d chars back", len(md))
            self.finished.emit(md, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("worker crashed")
            self.finished.emit("", err)


class RagKbWorker(QObject):
    """Runs one grounded RAG study prompt through the same backend as the GUI."""

    finished = Signal(str, str)
    progress_text = Signal(str)

    def __init__(
        self,
        system_prompt: str,
        user_msg: str,
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec: float = 240.0,
    ):
        super().__init__()
        self.system_prompt = system_prompt
        self.user_msg = user_msg
        self.model = model
        self.provider = provider
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.rag_worker")
        try:
            from quiz_ai_coach import _LOG_PATH, _run_llm_sync
            log.info("rag worker starting, prompt_chars=%d, log=%s", len(self.user_msg), _LOG_PATH)
            md = _run_llm_sync(
                self.system_prompt,
                self.user_msg,
                self.model,
                progress_cb=lambda s: self.progress_text.emit(s),
                timeout_sec=self.timeout_sec,
                provider=self.provider,
                max_output_tokens=8000,
            )
            log.info("rag worker got %d chars back", len(md))
            self.finished.emit(md, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("rag worker crashed")
            self.finished.emit("", err)


class MetaCoachWorker(QObject):
    """Runs quiz_ai_coach.analyze_meta_coach in a QThread."""
    finished = Signal(str, str)
    progress_text = Signal(str)

    def __init__(self, reports_md: str, model: Optional[str] = None, provider: str = "claude", timeout_sec: float = 600.0,
                 stats_payload: Optional[Dict] = None):
        super().__init__()
        self.reports_md = reports_md
        self.model = model
        self.provider = provider
        self.timeout_sec = timeout_sec
        self.stats_payload = stats_payload or {}

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.meta_worker")
        try:
            from quiz_ai_coach import analyze_meta_coach, _LOG_PATH
            log.info("meta worker starting, chars=%d, log=%s", len(self.reports_md), _LOG_PATH)
            md = analyze_meta_coach(
                self.reports_md,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
                stats_payload=self.stats_payload,
            )
            log.info("meta worker got %d chars back", len(md))
            self.finished.emit(md, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("meta worker crashed")
            self.finished.emit("", err)


class DeepReviewWorker(QObject):
    """Runs quiz_ai_coach.analyze_deep_review in a QThread."""
    finished = Signal(str, str)
    progress_text = Signal(str)
    one_done = Signal(int, str)  # (qid, markdown) - emitted as each parallel agent completes

    def __init__(
        self,
        items: List[Dict],
        prior_reports: Optional[Dict[int, str]] = None,
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec: float = 600.0,
    ):
        super().__init__()
        self.items = items
        self.prior_reports = prior_reports or {}
        self.model = model
        self.provider = provider
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.deep_worker")
        try:
            from quiz_ai_coach import analyze_deep_review, _LOG_PATH
            log.info(
                "deep worker starting, items=%d, with_prior=%d, log=%s",
                len(self.items), len(self.prior_reports), _LOG_PATH,
            )
            md = analyze_deep_review(
                self.items,
                prior_reports=self.prior_reports,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
                on_qid_done=lambda qid, md: self.one_done.emit(int(qid), md),
            )
            log.info("deep worker got %d chars back", len(md))
            self.finished.emit(md, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("deep worker crashed")
            self.finished.emit("", err)


class PreBriefWorker(QObject):
    """Runs quiz_ai_coach.analyze_pre_brief in a QThread (cheap context dossier)."""
    finished = Signal(str, str)
    progress_text = Signal(str)

    def __init__(
        self,
        items: List[Dict],
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec: float = 240.0,
    ):
        super().__init__()
        self.items = items
        self.model = model
        self.provider = provider
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.prebrief_worker")
        try:
            from quiz_ai_coach import analyze_pre_brief, _LOG_PATH
            log.info("prebrief worker starting, items=%d, log=%s", len(self.items), _LOG_PATH)
            md = analyze_pre_brief(
                self.items,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
            )
            log.info("prebrief worker got %d chars back", len(md))
            self.finished.emit(md, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("prebrief worker crashed")
            self.finished.emit("", err)


# Catalog of AWS networking concepts the dossier builder recognizes.
# Each entry: (canonical_name, slug, regex_pattern_for_detection).
# Patterns are case-insensitive word-boundary matches.
CONCEPT_CATALOG: List[tuple] = [
    ("Transit Gateway", "transit-gateway", r"\b(?:transit\s+gateway|tgw)\b"),
    ("Direct Connect", "direct-connect", r"\b(?:direct\s+connect|\bdx\b|private\s+vif|public\s+vif|transit\s+vif|macsec)\b"),
    ("VPC Peering", "vpc-peering", r"\bvpc\s+peering\b"),
    ("VPC Endpoints (Gateway + Interface)", "vpc-endpoints", r"\b(?:vpc\s+endpoint|gateway\s+endpoint|interface\s+endpoint|privatelink)\b"),
    ("Route 53 Resolver", "route53-resolver", r"\b(?:route\s*53\s+resolver|r53\s+resolver|inbound\s+endpoint|outbound\s+endpoint|resolver\s+rule)\b"),
    ("Route 53 DNS / Hosted Zones", "route53-dns", r"\b(?:hosted\s+zone|alias\s+record|delegation|route\s*53(?!\s+resolver))\b"),
    ("BGP", "bgp", r"\b(?:bgp|as[_\s-]?path|as[_\s-]?prepend|local[_\s-]?pref|med\b|bgp\s+communit)\b"),
    ("NAT Gateway / IGW / Egress-only IGW", "nat-igw", r"\b(?:nat\s+gateway|nat\s+gw|internet\s+gateway|\bigw\b|egress[_\s-]?only)\b"),
    ("Network Load Balancer", "nlb", r"\b(?:network\s+load\s+balancer|\bnlb\b)\b"),
    ("Application Load Balancer", "alb", r"\b(?:application\s+load\s+balancer|\balb\b)\b"),
    ("Gateway Load Balancer", "gwlb", r"\b(?:gateway\s+load\s+balancer|\bgwlb\b|geneve)\b"),
    ("Security Groups vs NACLs", "sg-nacl", r"\b(?:security\s+group|\bsg\b|nacl|network\s+acl|stateful|stateless)\b"),
    ("CloudFront", "cloudfront", r"\b(?:cloudfront|cf\s+distribution|origin\s+shield|edge\s+location)\b"),
    ("Site-to-Site VPN", "vpn-s2s", r"\b(?:site[_\s-]?to[_\s-]?site\s+vpn|s2s\s+vpn|customer\s+gateway|virtual\s+private\s+gateway|\bvgw\b)\b"),
    ("Client VPN", "client-vpn", r"\bclient\s+vpn\b"),
    ("WorkSpaces", "workspaces", r"\bworkspaces?\b"),
    ("VPC Flow Logs", "flow-logs", r"\b(?:flow\s+logs?|vpc\s+flow|pkt[_-]?(?:src|dst)[_-]?aws[_-]?service)\b"),
    ("IPAM / BYOIP", "ipam-byoip", r"\b(?:ipam|byoip|bring\s+your\s+own\s+ip)\b"),
    ("IMDS", "imds", r"\b(?:imds|instance\s+metadata|169\.254\.169\.254)\b"),
    ("ENI / EIP", "eni-eip", r"\b(?:elastic\s+network\s+interface|\beni\b|elastic\s+ip|\beip\b)\b"),
    ("Reachability Analyzer", "reachability-analyzer", r"\b(?:reachability\s+analyzer|network\s+access\s+analyzer)\b"),
    ("Cloud WAN / Network Manager", "cloud-wan", r"\b(?:cloud\s+wan|core\s+network|network\s+manager)\b"),
    ("AWS Global Accelerator", "global-accelerator", r"\b(?:global\s+accelerator|\bga\b\s+(?:listener|endpoint))\b"),
    ("Route Tables / Route Propagation", "route-tables", r"\b(?:route\s+table|route\s+propagation|propagated\s+route|longest\s+prefix\s+match|route\s+priority)\b"),
    ("EKS / VPC CNI", "eks-cni", r"\b(?:\beks\b|kubernetes|vpc\s+cni|secondary\s+eni\s+for\s+pods)\b"),
    ("Direct Connect Gateway", "dx-gateway", r"\b(?:direct\s+connect\s+gateway|\bdxgw\b|associated\s+gateway)\b"),
    ("Transit Gateway Connect", "tgw-connect", r"\b(?:tgw\s+connect|transit\s+gateway\s+connect|gre\s+tunnel)\b"),
]


def _detect_concepts(text: str) -> List[str]:
    """Return canonical concept names whose regex matches in the given text."""
    out: List[str] = []
    if not text:
        return out
    for name, _slug, pattern in CONCEPT_CATALOG:
        if re.search(pattern, text, flags=re.IGNORECASE):
            out.append(name)
    return out


def _slug_for_concept(name: str) -> str:
    for cn, slug, _ in CONCEPT_CATALOG:
        if cn == name:
            return slug
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


DOSSIER_MAX_WORKERS = 3
DOSSIER_MAX_EXCERPTS_PER_CONCEPT = 18
DOSSIER_MAX_CHARS_PER_EXCERPT = 2200


def _trim_dossier_excerpt_body(body: str) -> str:
    text = (body or "").strip()
    if len(text) <= DOSSIER_MAX_CHARS_PER_EXCERPT:
        return text
    cut = text[:DOSSIER_MAX_CHARS_PER_EXCERPT].rsplit("\n", 1)[0].strip()
    if len(cut) < DOSSIER_MAX_CHARS_PER_EXCERPT * 0.6:
        cut = text[:DOSSIER_MAX_CHARS_PER_EXCERPT].strip()
    return cut + "\n\n_[excerpt trimmed for dossier build]_"


def _pack_dossier_excerpts(excerpts: List[Dict[str, object]]) -> List[Dict[str, object]]:
    """Keep the dossier input focused: best unique qids, capped text per excerpt."""
    def _source_rank(src: object) -> int:
        return 0 if str(src or "") == "deep_review_reports.md" else 1

    def _rank(e: Dict[str, object]) -> tuple:
        body = str(e.get("body") or "")
        try:
            qid = int(e.get("qid") or 0)
        except Exception:
            qid = 0
        return (_source_rank(e.get("source")), -len(body), qid)

    best_by_qid: Dict[int, Dict[str, object]] = {}
    for e in excerpts:
        try:
            qid = int(e.get("qid") or 0)
        except Exception:
            qid = 0
        old = best_by_qid.get(qid)
        if old is None or _rank(e) < _rank(old):
            best_by_qid[qid] = e

    selected = sorted(best_by_qid.values(), key=_rank)[:DOSSIER_MAX_EXCERPTS_PER_CONCEPT]
    packed: List[Dict[str, object]] = []
    for e in selected:
        packed.append({
            "qid": e.get("qid"),
            "topic": e.get("topic") or "(no topic)",
            "body": _trim_dossier_excerpt_body(str(e.get("body") or "")),
            "source": e.get("source"),
        })
    return packed


# ---------------------------------------------------------------------------
# Persistent files in the project dir (alongside the .md report files)
#   srs_state.json       - SM-2-lite scheduling per qid
#   concept_mastery.json - right/wrong rollup per AWS concept
#   confidence_log.jsonl - append-only log: {ts, qid, confidence, correct, ans}
# ---------------------------------------------------------------------------

def _here(name: str) -> Path:
    return Path(__file__).with_name(name)


def _load_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _save_json_atomic(path: Path, data) -> None:
    try:
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)
    except Exception:
        pass


def _today_iso() -> str:
    return _date.today().isoformat()


def _srs_update(state: Dict[str, dict], qid: int, correct: bool) -> dict:
    """SM-2-lite. Returns the updated entry for this qid."""
    key = str(qid)
    e = state.get(key) or {"ease": 2.5, "interval": 0, "next_review": _today_iso(),
                           "last_reviewed": "", "reps": 0, "lapses": 0}
    ease = float(e.get("ease", 2.5) or 2.5)
    interval = int(e.get("interval", 0) or 0)
    reps = int(e.get("reps", 0) or 0)
    lapses = int(e.get("lapses", 0) or 0)
    if correct:
        reps += 1
        if reps == 1:
            interval = 1
        elif reps == 2:
            interval = 3
        else:
            interval = max(1, round(interval * ease))
        ease = min(2.8, ease + 0.05)
    else:
        reps = 0
        interval = 1
        lapses += 1
        ease = max(1.3, ease - 0.2)
    today = _date.today()
    nxt = (today + _timedelta(days=interval)).isoformat()
    e.update({"ease": round(ease, 2), "interval": interval, "next_review": nxt,
              "last_reviewed": today.isoformat(), "reps": reps, "lapses": lapses})
    state[key] = e
    return e


def _cm_bump(cm: Dict[str, dict], concepts: List[str], qid: int, correct: bool) -> None:
    today = _today_iso()
    for name in concepts:
        c = cm.get(name) or {"slug": _slug_for_concept(name), "right": 0, "wrong": 0,
                              "attempts": 0, "last_seen": "", "qids": []}
        c["attempts"] = int(c.get("attempts", 0)) + 1
        if correct:
            c["right"] = int(c.get("right", 0)) + 1
        else:
            c["wrong"] = int(c.get("wrong", 0)) + 1
        c["last_seen"] = today
        seen = set(int(x) for x in (c.get("qids") or []))
        seen.add(int(qid))
        c["qids"] = sorted(seen)
        cm[name] = c


def _conf_append(entry: dict) -> None:
    try:
        with _here("confidence_log.jsonl").open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


class ConceptDossierWorker(QObject):
    """Runs quiz_ai_coach.analyze_concept_dossier per concept in a small pool.

    Emits per-concept progress so the UI can show 'building X of N: <concept>'.
    """
    finished = Signal(list, str)  # list of (concept, slug, dossier_md), error
    progress_text = Signal(str)
    one_done = Signal(str, str, str)  # concept, slug, dossier_md (incremental save)

    def __init__(
        self,
        groups: List[Dict[str, Any]],
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec_per_concept: float = 360.0,
        max_workers: int = DOSSIER_MAX_WORKERS,
    ):
        super().__init__()
        self.groups = groups  # [{"concept": str, "excerpts": [{"qid", "topic", "body"}]}]
        self.model = model
        self.provider = provider
        self.timeout_sec_per_concept = timeout_sec_per_concept
        self.max_workers = max(1, int(max_workers or 1))
        self.cancel_requested = False

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
        log = logging.getLogger("quiz_ai_coach.dossier_worker")
        results: List[tuple] = []
        try:
            from quiz_ai_coach import analyze_concept_dossier
            groups = list(self.groups)
            total = len(groups)
            max_workers = min(self.max_workers, max(1, total))

            def _build_one(i: int, group: Dict[str, Any]) -> tuple:
                concept = group["concept"]
                excerpts = group["excerpts"]
                raw_count = int(group.get("raw_excerpt_count") or len(excerpts))
                slug = _slug_for_concept(concept)
                self.progress_text.emit(
                    f"[{i}/{total}] {concept} - {len(excerpts)}/{raw_count} excerpts"
                )
                log.info(
                    "dossier %d/%d concept=%r excerpts=%d raw=%d",
                    i, total, concept, len(excerpts), raw_count,
                )
                md = analyze_concept_dossier(
                    concept, excerpts,
                    model=self.model,
                    provider=self.provider,
                    timeout_sec=self.timeout_sec_per_concept,
                    progress_cb=lambda s, c=concept: self.progress_text.emit(f"{c}: {s}"),
                )
                return (i, concept, slug, md)

            self.progress_text.emit(f"Starting {total} dossiers with {max_workers} parallel workers")
            pending = {}
            next_index = 0
            with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="dossier") as pool:
                def _submit_next() -> None:
                    nonlocal next_index
                    if self.cancel_requested or next_index >= total:
                        return
                    next_index += 1
                    group = groups[next_index - 1]
                    pending[pool.submit(_build_one, next_index, group)] = next_index

                for _ in range(max_workers):
                    _submit_next()

                while pending:
                    done, _waiting = wait(set(pending.keys()), return_when=FIRST_COMPLETED)
                    for fut in done:
                        pending.pop(fut, None)
                        if self.cancel_requested:
                            continue
                        i, concept, slug, md = fut.result()
                        results.append((i, concept, slug, md))
                        self.one_done.emit(concept, slug, md)
                        _submit_next()

            ordered = [(concept, slug, md) for _i, concept, slug, md in sorted(results)]
            self.finished.emit(ordered, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("dossier worker crashed")
            ordered = [(concept, slug, md) for _i, concept, slug, md in sorted(results)]
            self.finished.emit(ordered, err)


class NuclearReviewWorker(QObject):
    """Runs quiz_ai_coach.analyze_nuclear_review (multi-agent fan-out) for ONE qid."""
    finished = Signal(dict, str)  # {final, boundary, patterns, distractors}, error
    progress_text = Signal(str)

    def __init__(
        self,
        item: Dict,
        master_excerpts: List[Dict],
        dossier_md: str,
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec: float = 600.0,
    ):
        super().__init__()
        self.item = item
        self.master_excerpts = master_excerpts
        self.dossier_md = dossier_md
        self.model = model
        self.provider = provider
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.nuclear_worker")
        try:
            from quiz_ai_coach import analyze_nuclear_review
            log.info(
                "nuclear worker starting, qid=%s, history=%d, dossier=%d chars",
                self.item.get("qid"), len(self.master_excerpts), len(self.dossier_md or ""),
            )
            result = analyze_nuclear_review(
                self.item,
                master_excerpts=self.master_excerpts,
                dossier_md=self.dossier_md,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
            )
            log.info("nuclear worker got final=%d chars", len(result.get("final", "")))
            self.finished.emit(result, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("nuclear worker crashed")
            self.finished.emit({}, err)


class DiagramWorker(QObject):
    """Runs quiz_ai_coach.analyze_diagram for ONE qid. Returns raw HTML."""
    finished = Signal(str, str)        # (html, error_or_empty)
    progress_text = Signal(str)

    # Sonnet 4.6 produced visually broken layouts (lifeline-skipping arrows,
    # diagonal lines, labels overlapping boxes) even with ~30 layout rules in
    # the prompt; it prioritized completion over rule compliance. Opus 4.7
    # with adaptive thinking + xhigh effort respects the spatial constraints.
    DEFAULT_MODEL = "claude-opus-4-7"

    def __init__(
        self,
        item: Dict,
        prior_artifacts: Optional[Dict[str, str]] = None,
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec: float = 600.0,
    ):
        super().__init__()
        self.item = item
        self.prior_artifacts = prior_artifacts or {}
        self.provider = provider
        self.model = model or (self.DEFAULT_MODEL if provider == "claude" else None)
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.diagram_worker")
        try:
            from quiz_ai_coach import analyze_diagram
            log.info(
                "diagram worker starting, qid=%s, prior_keys=%s",
                self.item.get("qid"), sorted(self.prior_artifacts.keys()),
            )
            html_str = analyze_diagram(
                self.item,
                prior_artifacts=self.prior_artifacts,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
            )
            log.info("diagram worker got %d chars back", len(html_str))
            self.finished.emit(html_str, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("diagram worker crashed")
            self.finished.emit("", err)


class TeachZeroWorker(QObject):
    """Runs quiz_ai_coach.analyze_teach_zero for ONE qid. Returns markdown."""
    finished = Signal(str, str)        # (markdown, error_or_empty)
    progress_text = Signal(str)

    DEFAULT_MODEL = "claude-opus-4-7"

    def __init__(
        self,
        item: Dict,
        prior_artifacts: Optional[Dict[str, str]] = None,
        model: Optional[str] = None,
        provider: str = "claude",
        timeout_sec: float = 600.0,
    ):
        super().__init__()
        self.item = item
        self.prior_artifacts = prior_artifacts or {}
        self.provider = provider
        self.model = model or (self.DEFAULT_MODEL if provider == "claude" else None)
        self.timeout_sec = timeout_sec

    @Slot()
    def run(self) -> None:
        import logging, traceback as _tb
        log = logging.getLogger("quiz_ai_coach.teach_zero_worker")
        try:
            from quiz_ai_coach import analyze_teach_zero
            log.info(
                "teach_zero worker starting, qid=%s, prior_keys=%s",
                self.item.get("qid"), sorted(self.prior_artifacts.keys()),
            )
            md_str = analyze_teach_zero(
                self.item,
                prior_artifacts=self.prior_artifacts,
                model=self.model,
                provider=self.provider,
                timeout_sec=self.timeout_sec,
                progress_cb=lambda s: self.progress_text.emit(s),
            )
            log.info("teach_zero worker got %d chars back", len(md_str))
            self.finished.emit(md_str, "")
        except Exception as e:
            err = f"{type(e).__name__}: {e}\n\n{_tb.format_exc()}"
            log.exception("teach_zero worker crashed")
            self.finished.emit("", err)


# -----------------------------
# UI helpers
# -----------------------------
def _fmt_mmss(seconds: int) -> str:
    if seconds < 0:
        seconds = 0
    m = seconds // 60
    s = seconds % 60
    return f"{m:02d}:{s:02d}"


def _status_icon(correct: Optional[bool], flagged: bool, answered: bool) -> str:
    if correct is True:
        return "\u2705"
    if correct is False:
        return "\u274c"
    if flagged:
        return "\u2691"
    if answered:
        return "\u25cf"
    return "\u25cb"


class OptionRow(QWidget):
    """
    A row that wraps long option text nicely:
    [control]  (A.) <wrapped label>
    Clicking the label toggles the control.
    """
    toggled = Signal(str, bool)  # (letter, checked)

    def __init__(self, letter: str, text: str, is_multi: bool):
        super().__init__()
        self.letter = letter
        self.is_multi = is_multi

        lay = QHBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(10)

        if is_multi:
            self.ctrl = QCheckBox()
            self.ctrl.stateChanged.connect(self._emit)
        else:
            self.ctrl = QRadioButton()
            self.ctrl.toggled.connect(lambda checked: self._emit(checked))

        self.lbl = QLabel(f"<b>{letter}.</b> {text}")
        self.lbl.setTextFormat(Qt.RichText)
        self.lbl.setWordWrap(True)
        self.lbl.setCursor(QCursor(Qt.PointingHandCursor))
        self.lbl.mousePressEvent = self._label_clicked  # type: ignore

        lay.addWidget(self.ctrl, 0, Qt.AlignTop)
        lay.addWidget(self.lbl, 1)

        self.setObjectName("OptionRow")
        self.setStyleSheet("""
            QWidget#OptionRow {
                border: 1px solid palette(mid);
                border-radius: 10px;
            }
            QWidget#OptionRow:hover {
                border: 1px solid palette(dark);
            }
        """)

    def _label_clicked(self, _ev):
        if self.is_multi:
            self.ctrl.setChecked(not self.ctrl.isChecked())
        else:
            self.ctrl.setChecked(True)

    def set_checked(self, checked: bool):
        self.ctrl.blockSignals(True)
        self.ctrl.setChecked(checked)
        self.ctrl.blockSignals(False)

    def is_checked(self) -> bool:
        return bool(self.ctrl.isChecked())

    def _emit(self, *args):
        checked = self.is_checked()
        self.toggled.emit(self.letter, checked)


class ResultsDialog(QDialog):
    def __init__(self, parent: QWidget, total: int, correct: int, unanswered: int, seconds_spent: int):
        super().__init__(parent)
        self.setWindowTitle("Results")
        self.setMinimumWidth(440)

        pct = (correct / total) * 100 if total else 0.0
        mins = seconds_spent // 60
        secs = seconds_spent % 60

        box = QVBoxLayout(self)
        title = QLabel("Exam results")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        box.addWidget(title)

        lines = [
            f"Questions: {total}",
            f"Correct: {correct}",
            f"Unanswered: {unanswered}",
            f"Score: {pct:.1f}%",
            f"Time: {mins}m {secs}s",
        ]
        for t in lines:
            lab = QLabel(t)
            lab.setTextInteractionFlags(Qt.TextSelectableByMouse)
            box.addWidget(lab)

        btn = QPushButton("Close")
        btn.clicked.connect(self.accept)
        box.addSpacing(12)
        box.addWidget(btn, alignment=Qt.AlignRight)


# -----------------------------
# Main Window
# -----------------------------


class StatsDialog(QDialog):
    def __init__(
        self,
        parent: QWidget,
        q_stats: Dict[int, Dict[str, int]],
        p_stats: Dict[str, Dict[str, int]],
        t_stats: Dict[str, Dict[str, int]],
        bank_pattern_counts: Dict[str, int],
        bank_tag_counts: Dict[str, int],
    ):
        super().__init__(parent)
        self.setWindowTitle("Stats")
        self.setMinimumSize(720, 560)

        tabs = QTabWidget()
        lay = QVBoxLayout(self)
        lay.addWidget(tabs, 1)

        def mk_table(rows: List[Tuple[str, int, int, int, float]]) -> str:
            # rows: name, attempts, wrong, correct, wrong_rate
            out = []
            out.append("<table border='0' cellpadding='6' cellspacing='0' style='width:100%'>")
            out.append(
                "<tr style='opacity:0.85'>"
                "<th align='left'>Item</th>"
                "<th align='right'>Attempts</th>"
                "<th align='right'>Wrong</th>"
                "<th align='right'>Correct</th>"
                "<th align='right'>Wrong%</th>"
                "</tr>"
            )
            for name, a, w, c, wr in rows:
                out.append(
                    "<tr>"
                    f"<td><b>{name}</b></td>"
                    f"<td align='right'>{a}</td>"
                    f"<td align='right'>{w}</td>"
                    f"<td align='right'>{c}</td>"
                    f"<td align='right'>{wr:.0f}%</td>"
                    "</tr>"
                )
            out.append("</table>")
            return "".join(out)

        def rows_from_agg(d: Dict[str, Dict[str, int]]) -> List[Tuple[str, int, int, int, float]]:
            rows: List[Tuple[str, int, int, int, float]] = []
            for name, s in d.items():
                a = int(s.get("attempts", 0) or 0)
                w = int(s.get("wrong", 0) or 0)
                c = int(s.get("correct", 0) or 0)
                wr = (100.0 * w / a) if a else 0.0
                rows.append((name, a, w, c, wr))
            return rows

        # ---- Questions tab ----
        q_view = QTextBrowser()
        q_view.setOpenExternalLinks(False)
        q_rows = []
        for qid, s in q_stats.items():
            a = int(s.get("attempts", 0) or 0)
            w = int(s.get("wrong", 0) or 0)
            c = int(s.get("correct", 0) or 0)
            if a <= 0:
                continue
            wr = 100.0 * w / a
            q_rows.append((f"Q{qid}", a, w, c, wr))
        q_rows.sort(key=lambda r: (r[2], r[4], r[1]), reverse=True)
        q_view.setHtml(
            "<h2>Mistakes per question</h2>"
            "<div style='opacity:0.85'>Sorted by wrong count (then wrong%).</div>"
            + mk_table(q_rows)
        )
        tabs.addTab(q_view, "Questions")

        # ---- Patterns tab ----
        p_view = QTextBrowser()
        p_rows = rows_from_agg(p_stats)
        p_rows.sort(key=lambda r: (r[2], r[4], r[1]), reverse=True)

        # top by wrong rate (require some attempts so it isn't noisy)
        p_top_rate = [r for r in p_rows if r[1] >= 3]
        p_top_rate.sort(key=lambda r: (r[4], r[2], r[1]), reverse=True)

        # wrong bank distribution
        bp = sorted(bank_pattern_counts.items(), key=lambda kv: kv[1], reverse=True)
        bp_rows = [(name, cnt, cnt, 0, 0.0) for name, cnt in bp]  # reuse table formatter

        p_view.setHtml(
            "<h2>Patterns</h2>"
            "<h3>Top patterns you miss most (Wrong%)</h3>"
            + (mk_table(p_top_rate[:20]) if p_top_rate else "<div style='opacity:0.85'>No pattern stats yet.</div>")
            + "<hr>"
            "<h3>Most missed patterns (Wrong count)</h3>"
            + (mk_table(p_rows[:50]) if p_rows else "<div style='opacity:0.85'>No pattern stats yet.</div>")
            + "<hr>"
            "<h3>Patterns inside your Wrong DB</h3>"
            + (mk_table(bp_rows[:50]) if bp_rows else "<div style='opacity:0.85'>Wrong DB is empty (or no labels yet).</div>")
        )
        tabs.addTab(p_view, "Patterns")

        # ---- Tags tab ----
        t_view = QTextBrowser()
        t_rows = rows_from_agg(t_stats)
        t_rows.sort(key=lambda r: (r[2], r[4], r[1]), reverse=True)

        bt = sorted(bank_tag_counts.items(), key=lambda kv: kv[1], reverse=True)
        bt_rows = [(name, cnt, cnt, 0, 0.0) for name, cnt in bt]

        t_top_rate = [r for r in t_rows if r[1] >= 3]
        t_top_rate.sort(key=lambda r: (r[4], r[2], r[1]), reverse=True)

        t_view.setHtml(
            "<h2>Tags</h2>"
            "<h3>Top tags you miss most (Wrong%)</h3>"
            + (mk_table(t_top_rate[:25]) if t_top_rate else "<div style='opacity:0.85'>No tag stats yet.</div>")
            + "<hr>"
            "<h3>Most missed tags (Wrong count)</h3>"
            + (mk_table(t_rows[:60]) if t_rows else "<div style='opacity:0.85'>No tag stats yet.</div>")
            + "<hr>"
            "<h3>Tags inside your Wrong DB</h3>"
            + (mk_table(bt_rows[:60]) if bt_rows else "<div style='opacity:0.85'>Wrong DB is empty (or no labels yet).</div>")
        )
        tabs.addTab(t_view, "Tags")
class DeepReviewPickerDialog(QDialog):
    """Picker shared by Deep Review and Pre-Brief.

    Scopes are richer than AI Coach: big offenders (most-failed first), repeat,
    current, bank, custom qid list, no-prior-AI-Coach. Each shows total / with
    prior AI Coach / already deep-reviewed.
    """
    def __init__(
        self,
        parent: QWidget,
        *,
        scopes: Dict[str, Dict[str, int]],
        repeat_threshold: int,
        big_threshold: int,
        wrong_bank_total: int,
        batch_size_default: int,
        mode_label: str = "Deep Review",
    ):
        super().__init__(parent)
        self.setWindowTitle(f"{mode_label} - pick scope")
        self.resize(720, 520)
        self.choice: Optional[str] = None
        self.batch_size: int = batch_size_default
        self.custom_qids_text: str = ""
        self.force_redo: bool = False
        self._scopes = scopes

        from PySide6.QtWidgets import QRadioButton as _RB, QSpinBox, QGridLayout, QLineEdit, QCheckBox

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"<b>{mode_label} - which qids?</b>"
        ))

        def _row(key: str, label: str) -> "_RB":
            s = scopes.get(key, {"total": 0, "with_prior": 0, "deep_done": 0, "pending": 0})
            rb = _RB(
                f"{label}  -  {s['pending']} pending"
                + (f" ({s.get('deep_done', 0)} already done)" if s.get('deep_done') else "")
            )
            rb.setEnabled(s['total'] > 0)
            return rb

        self._rb_big = _row("big", f"\u2b50 Big offenders (wrong >= {big_threshold}, sorted desc)")
        self._rb_repeat = _row("repeat", f"Repeat offenders (wrong >= {repeat_threshold})")
        self._rb_current = _row("current", "Current session (live wrong)")
        self._rb_bank = _row("bank", f"Wrong DB (saved + AI/Deep, {wrong_bank_total} qids)")
        self._rb_no_prior = _row("no_prior", "Sin AI Coach previo (huerfanos del round actual)")
        self._rb_custom = _RB("Custom qid list (comma/space separated)")

        # default = first scope with pending > 0
        for rb, key in (
            (self._rb_big, "big"),
            (self._rb_repeat, "repeat"),
            (self._rb_current, "current"),
            (self._rb_bank, "bank"),
            (self._rb_no_prior, "no_prior"),
        ):
            if scopes.get(key, {}).get("pending", 0) > 0:
                rb.setChecked(True)
                break
        else:
            self._rb_big.setChecked(True)

        for rb in (self._rb_big, self._rb_repeat, self._rb_current, self._rb_bank, self._rb_no_prior, self._rb_custom):
            layout.addWidget(rb)

        custom_row = QHBoxLayout()
        custom_row.addWidget(QLabel("    qids:"))
        self._custom_edit = QLineEdit()
        self._custom_edit.setPlaceholderText("e.g.  12, 45 99 230")
        custom_row.addWidget(self._custom_edit, 1)
        layout.addLayout(custom_row)

        cfg = QGridLayout()
        cfg.addWidget(QLabel("Batch size:"), 0, 0)
        self._spin = QSpinBox()
        self._spin.setRange(1, 20)
        self._spin.setValue(batch_size_default)
        self._spin.setSuffix(" qids per call")
        cfg.addWidget(self._spin, 0, 1)
        cfg.setColumnStretch(2, 1)
        layout.addLayout(cfg)

        self._chk_force = QCheckBox("Re-analyze even if already on disk (quema tokens)")
        layout.addWidget(self._chk_force)

        layout.addWidget(QLabel(
            "<span style='opacity:0.7'>Tip: Deep Review es caro (~2-3 min/qid). Empeza con "
            "batch 3 sobre Big offenders. Pre-Brief es barato y prepara terreno.</span>"
        ))

        btn_row = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_ok = QPushButton(f"Run {mode_label}")
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(self.reject)
        btn_ok.clicked.connect(self._accept)
        btn_row.addStretch(1)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _selected_scope(self) -> str:
        for rb, key in (
            (self._rb_big, "big"),
            (self._rb_repeat, "repeat"),
            (self._rb_current, "current"),
            (self._rb_bank, "bank"),
            (self._rb_no_prior, "no_prior"),
            (self._rb_custom, "custom"),
        ):
            if rb.isChecked():
                return key
        return "big"

    def _accept(self) -> None:
        self.choice = self._selected_scope()
        self.batch_size = int(self._spin.value())
        self.custom_qids_text = self._custom_edit.text().strip()
        self.force_redo = bool(self._chk_force.isChecked())
        if self.choice == "custom" and not self.custom_qids_text:
            QMessageBox.information(self, "Custom qids", "Type at least one qid (e.g. 12, 45).")
            return
        self.accept()


class AICoachPickerDialog(QDialog):
    """Asks the user which set of wrong questions to send to the AI Coach.

    Shows per-scope progress (total / already analyzed / pending) and lets the
    user pick batch size + reset progress per scope.
    """
    def __init__(
        self,
        parent: QWidget,
        *,
        scopes: Dict[str, Dict[str, int]],
        repeat_threshold: int,
        wrong_bank_total: int,
        batch_size_default: int,
    ):
        super().__init__(parent)
        self.setWindowTitle("AI Coach - pick what to analyze")
        self.resize(640, 420)
        self.choice: Optional[str] = None  # "current" | "bank" | "repeat"
        self.batch_size: int = batch_size_default
        self.reset_requested: bool = False
        self._scopes = scopes

        from PySide6.QtWidgets import QRadioButton as _RB, QSpinBox, QGridLayout

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "<b>Which wrong answers should the AI Coach analyze?</b><br>"
            "<span style='opacity:0.8'>Each click sends only the next batch - progress is "
            "remembered per scope so you can keep clicking until done.</span>"
        ))

        def _row(key: str, label: str) -> "_RB":
            s = scopes[key]
            rb = _RB(
                f"{label}  -  {s['pending']} pending"
                + (f" ({s['sent']} already analyzed)" if s['sent'] else "")
            )
            rb.setEnabled(s['total'] > 0)
            return rb

        self._rb_current = _row("current", "Current session (live wrong)")
        self._rb_bank = _row("bank", f"Wrong DB (saved + AI/Deep, {wrong_bank_total} qids)")
        self._rb_repeat = _row("repeat", f"Repeat offenders (wrong >= {repeat_threshold})")

        # default = first scope with pending > 0
        for rb, key in ((self._rb_current, "current"), (self._rb_bank, "bank"), (self._rb_repeat, "repeat")):
            if scopes[key]["pending"] > 0:
                rb.setChecked(True)
                break
        else:
            self._rb_current.setChecked(True)

        layout.addWidget(self._rb_current)
        layout.addWidget(self._rb_bank)
        layout.addWidget(self._rb_repeat)

        cfg = QGridLayout()
        cfg.addWidget(QLabel("Batch size:"), 0, 0)
        self._spin = QSpinBox()
        self._spin.setRange(1, 50)
        self._spin.setValue(batch_size_default)
        self._spin.setSuffix(" questions per call")
        cfg.addWidget(self._spin, 0, 1)
        cfg.setColumnStretch(2, 1)
        layout.addLayout(cfg)

        layout.addWidget(QLabel(
            "<span style='opacity:0.7'>Tip: 10 keeps each Opus call fast and focused. "
            "Increase if you want fewer round-trips.</span>"
        ))

        btn_row = QHBoxLayout()
        btn_reset = QPushButton("Reset progress (selected scope)")
        btn_reset.clicked.connect(self._reset)
        btn_cancel = QPushButton("Cancel")
        btn_ok = QPushButton("Analyze next batch")
        btn_ok.setDefault(True)
        btn_cancel.clicked.connect(self.reject)
        btn_ok.clicked.connect(self._accept)
        btn_row.addWidget(btn_reset)
        btn_row.addStretch(1)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _selected_scope(self) -> str:
        if self._rb_current.isChecked():
            return "current"
        if self._rb_bank.isChecked():
            return "bank"
        return "repeat"

    def _accept(self) -> None:
        self.choice = self._selected_scope()
        self.batch_size = int(self._spin.value())
        self.accept()

    def _reset(self) -> None:
        self.choice = self._selected_scope()
        self.reset_requested = True
        self.batch_size = int(self._spin.value())
        self.accept()


def _style_ai_coach_text_view(view: QTextBrowser, *, raw: bool = False) -> None:
    palette = view.palette()
    text = palette.color(QPalette.ColorRole.Text).name()
    base = palette.color(QPalette.ColorRole.Base).name()
    alt = palette.color(QPalette.ColorRole.AlternateBase).name()
    link = palette.color(QPalette.ColorRole.Link).name()
    mid = palette.color(QPalette.ColorRole.Mid).name()

    view.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
    view.document().setDocumentMargin(16 if not raw else 12)
    view.setStyleSheet("""
        QTextBrowser#AICoachReport,
        QTextBrowser#AICoachRaw {
            border: 1px solid palette(mid);
            border-radius: 6px;
            padding: 0;
        }
    """)

    if raw:
        view.setFont(QFont("Cascadia Mono", 10))
        return

    view.document().setDefaultFont(QFont("Segoe UI", 10))
    view.document().setDefaultStyleSheet(f"""
        body {{
            color: {text};
            background-color: {base};
            font-family: 'Segoe UI', sans-serif;
            font-size: 10.5pt;
            line-height: 145%;
        }}
        h1, h2, h3 {{
            color: {link};
            font-weight: 700;
            margin-top: 18px;
            margin-bottom: 8px;
        }}
        h3 {{
            font-size: 14pt;
            padding-bottom: 6px;
            border-bottom: 1px solid {mid};
        }}
        p {{
            margin-top: 6px;
            margin-bottom: 10px;
        }}
        ul, ol {{
            margin-top: 4px;
            margin-bottom: 12px;
        }}
        li {{
            margin-top: 5px;
            margin-bottom: 5px;
        }}
        strong {{
            color: {text};
            font-weight: 700;
        }}
        code {{
            color: {text};
            background-color: {alt};
            font-family: 'Cascadia Mono', Consolas, monospace;
            font-size: 9.5pt;
        }}
        pre {{
            color: {text};
            background-color: {alt};
            border: 1px solid {mid};
            padding: 8px;
            white-space: pre-wrap;
        }}
        a {{
            color: {link};
            text-decoration: none;
        }}
    """)


class AICoachDialog(QDialog):
    """Renders the AI coach markdown report inside a tabbed dialog."""
    def __init__(self, parent: QWidget, report_md: str, num_questions: int, model: str,
                 master_path: Optional[Path] = None, header_text: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("AI Coach - Why you missed those")
        self.resize(1100, 740)
        self._master_path = master_path

        # Parse qid index: Q<N> headings. AI Coach uses "### Q<N>";
        # Nuclear Review uses "# Q<N>" inside each saved run.
        # so clicking the index renders ONE section, not the whole 800k-char doc
        # (QTextBrowser.setMarkdown breaks past ~200k chars).
        self._qid_index: List[Tuple[int, str]] = []
        self._qid_sections: Dict[int, str] = {}
        heading_iter = list(re.finditer(r"^#{1,6}\s+Q(\d+)\b[^\n]*$", report_md, re.MULTILINE))
        for i, m in enumerate(heading_iter):
            qid = int(m.group(1))
            line = m.group(0)
            topic = re.sub(r"^#{1,6}\s+Q\d+\b\s*", "", line).strip()
            topic = re.sub(r"^(?:[-\u2013\u2014]\s*)+", "", topic).strip()[:60]
            if not topic:
                topic = "(untitled)"
            start = m.start()
            end = heading_iter[i + 1].start() if i + 1 < len(heading_iter) else len(report_md)
            section = report_md[start:end].rstrip()
            if qid in self._qid_sections:
                # Multiple analyses of same qid - concatenate, separated by hr
                self._qid_sections[qid] = self._qid_sections[qid] + "\n\n---\n\n" + section
            else:
                self._qid_sections[qid] = section
                self._qid_index.append((qid, topic))
        self._full_report_md = report_md
        # Header (everything before the first ### Q heading) - always shown on landing.
        self._report_header = report_md[:heading_iter[0].start()] if heading_iter else report_md

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        if header_text:
            header = QLabel(
                f"<b>{html.escape(header_text)}</b> "
                f"&nbsp;|&nbsp; model: <code>{html.escape(model or 'default')}</code>"
            )
        else:
            header = QLabel(
                f"<b>Analyzed {num_questions} wrong answer(s)</b> "
                f"&nbsp;|&nbsp; model: <code>{html.escape(model or 'default')}</code>"
            )
        header.setTextFormat(Qt.TextFormat.RichText)
        header.setWordWrap(False)
        header.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        header.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        header.setMinimumHeight(30)
        header.setMaximumHeight(34)
        header.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        header.setObjectName("AICoachHeader")
        header.setStyleSheet("""
            QLabel#AICoachHeader {
                border: 1px solid palette(mid);
                border-radius: 6px;
                padding: 4px 10px;
                background: palette(alternate-base);
            }
        """)
        layout.addWidget(header)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # â”€â”€ Left panel: search + index â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 6, 0)
        left_layout.setSpacing(6)
        left.setMinimumWidth(230)
        left.setMaximumWidth(330)

        self._search = QLineEdit()
        self._search.setPlaceholderText("Search Q# or topic...")
        self._search.setClearButtonEnabled(True)
        self._search.setStyleSheet("padding: 6px 8px; border-radius: 6px;")
        left_layout.addWidget(self._search)

        self._index_list = QListWidget()
        self._index_list.setAlternatingRowColors(True)
        self._index_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._index_list.setStyleSheet("""
            QListWidget {
                border: 1px solid palette(mid);
                border-radius: 6px;
                padding: 4px;
            }
            QListWidget::item {
                padding: 7px 8px;
                border-radius: 4px;
            }
            QListWidget::item:selected {
                background: palette(highlight);
                color: palette(highlighted-text);
            }
        """)
        for qid, topic in self._qid_index:
            item = QListWidgetItem(f"Q{qid}  {topic}")
            item.setData(Qt.ItemDataRole.UserRole, qid)
            self._index_list.addItem(item)
        left_layout.addWidget(self._index_list)
        splitter.addWidget(left)

        # â”€â”€ Right panel: tabs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        tabs = QTabWidget()

        self._view = QTextBrowser()
        self._view.setOpenLinks(False)
        self._view.setOpenExternalLinks(False)
        self._view.setObjectName("AICoachReport")
        _style_ai_coach_text_view(self._view)
        self._view.anchorClicked.connect(self._on_anchor_clicked)
        # Initial view: just the report header (overview stats), not the whole
        # archive. Clicking an index entry loads a single section. Avoids the
        # QTextBrowser markdown-renderer breakdown on 800k-char inputs.
        landing = self._report_header.strip() or "_(select a question on the left to view its analysis)_"
        if self._qid_index:
            landing += "\n\n---\n\n_Select a question on the left to view its analysis._"
        try:
            self._view.setMarkdown(self._linkify(landing))
        except Exception:
            self._view.setPlainText(landing)
        tabs.addTab(self._view, "Report")

        raw = QTextBrowser()
        raw.setObjectName("AICoachRaw")
        _style_ai_coach_text_view(raw, raw=True)
        raw.setPlainText(report_md)
        tabs.addTab(raw, "Raw markdown")

        splitter.addWidget(tabs)
        splitter.setSizes([260, 840])
        splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(splitter, 1)

        btn_row = QHBoxLayout()
        btn_save = QPushButton("Save report...")
        btn_save.clicked.connect(lambda: self._save(report_md))
        btn_open_all = QPushButton("Open all reports")
        btn_open_all.setEnabled(bool(master_path and master_path.exists()))
        btn_open_all.clicked.connect(self._open_master)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_open_all)
        btn_row.addStretch(1)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

        # Wire signals after all widgets exist
        self._index_list.itemClicked.connect(self._on_index_clicked)
        self._search.textChanged.connect(self._on_search_changed)

    def _on_anchor_clicked(self, url: QUrl) -> None:
        s = url.toString()
        if s.startswith("qid:"):
            try:
                qid = int(s.split(":", 1)[1])
            except ValueError:
                return
            parent = self.parent()
            jumped = False
            if parent is not None and hasattr(parent, "jump_to_qid"):
                jumped = bool(parent.jump_to_qid(qid))
                try:
                    parent.raise_()
                    parent.activateWindow()
                except Exception:
                    pass
                self.raise_()
                self.activateWindow()
            if not jumped:
                QMessageBox.information(
                    self, "Jump to question",
                    f"Q{qid} is not in the current session order.\n\n"
                    "Load the bank/scope that contains it and try again."
                )
            return
        if s:
            QDesktopServices.openUrl(url)

    def _linkify(self, md: str) -> str:
        """Wrap Q### tokens as clickable qid links so anchorClicked can jump."""
        return re.sub(r"\bQ(\d+)\b", r"[Q\1](qid:\1)", md)

    def _on_index_clicked(self, item: "QListWidgetItem") -> None:
        qid = item.data(Qt.ItemDataRole.UserRole)
        if qid is None:
            return
        section = self._qid_sections.get(int(qid))
        if not section:
            return
        try:
            self._view.setMarkdown(self._linkify(section))
        except Exception:
            self._view.setPlainText(section)
        # scroll to top of the freshly-rendered section
        cursor = self._view.textCursor()
        cursor.movePosition(cursor.MoveOperation.Start)
        self._view.setTextCursor(cursor)
        self._view.ensureCursorVisible()

    def _on_search_changed(self, text: str) -> None:
        text = text.strip().lower()
        for i in range(self._index_list.count()):
            item = self._index_list.item(i)
            item.setHidden(bool(text) and text not in item.text().lower())

    def _open_master(self) -> None:
        if not self._master_path:
            return
        try:
            os.startfile(str(self._master_path))  # type: ignore[attr-defined]
        except Exception as e:
            QMessageBox.warning(self, "Open failed", f"{e}\n\nPath: {self._master_path}")

    def _save(self, md: str) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save AI Coach report",
            f"ai_coach_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
            "Markdown (*.md);;All files (*.*)",
        )
        if path:
            try:
                Path(path).write_text(md, encoding="utf-8")
            except Exception as e:
                QMessageBox.warning(self, "Save failed", str(e))


class QuizWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ANS-C01 Quiz GUI")
        self.setMinimumSize(1180, 740)
        self.resize(1400, 860)
        try:
            self.setWindowIcon(QIcon.fromTheme("help-about"))
        except Exception:
            pass

        self.settings = QSettings("MGBTools", "ANSQuiz")
        self._session_profile_names_key = "session/profile_names"
        self._session_autosave_key = "session/autosave"
        self._session_profile_selected_key = "session/profile_selected"
        self._autosave_profile_name = "saved"
        self._pending_restore_state: Optional[Dict[str, object]] = None
        self._pending_restore_source: str = ""

        # Bank/state
        self.docx_path: Optional[Path] = None
        self.overrides_path: Optional[Path] = None
        self.overrides_stamp: Optional[Tuple[int, int]] = None
        self.questions: List[Question] = []
        self.docx_signature: str = ""
        self.bank_signature: str = ""

        # Session-only NEW tracking (questions with 0 historical attempts at session start)
        self.session_new_initial: Set[int] = set()
        self.session_new_remaining: Set[int] = set()


        self.overrides: Dict[int, str] = {}
        self.notes: Dict[int, str] = {}
        self.my_expl: Dict[int, str] = {}
        self.ov_patterns: Dict[int, str] = {}
        self.ov_tags: Dict[int, List[str]] = {}

        self.order: List[int] = []  # question indices in current session
        self.idx: int = 0
        self.selections: Dict[int, Set[str]] = {}
        self.flagged: Set[int] = set(self._load_flagged_bank())
        self.submitted: bool = False

        self.mode: str = "Exam"
        self.time_limit_min: int = 170

        # session sizing: 0 = all questions in the selected bank
        self.session_count: int = int(self.settings.value('quiz/count', 0) or 0)

        # Wrong filter (minimum wrong count for 'Wrong >= Min')
        try:
            self.min_wrong: int = max(0, int(self.settings.value('quiz/min_wrong', 2)))
        except Exception:
            self.min_wrong = 2

        # Custom question IDs (comma-separated). If set, overrides bank selection.
        self.custom_ids_text: str = self.settings.value('quiz/custom_ids', '', str) or ''
        self.custom_ids_active: bool = False

        self.shuffle_questions: bool = True
        self.shuffle_options: bool = True

        # Practice QoL: when enabled, Next auto-checks (does not affect persistent wrong stats)
        self.auto_check_next_enabled: bool = bool(int(self.settings.value('quiz/auto_check_next', 0) or 0))

        # last checked selection in practice (for 'Next checks once, next advances')
        self._last_checked_sel: Dict[int, FrozenSet[str]] = {}

        # Prevent double-counting stats on repeated Finish/Submit
        self._session_results_recorded: bool = False

        # option order per qid for this session (stable randomization)
        self.option_order: Dict[int, List[str]] = {}
        self._rng = random.Random()

        # keep radio group alive between renders (prevents leaks)
        self._option_group: Optional[QButtonGroup] = None

        # timing
        self.seconds_left: Optional[int] = None
        self.seconds_spent: int = 0
        self._tick = QTimer(self)
        self._tick.timeout.connect(self._on_tick)
        self._autosave_debounce = QTimer(self)
        self._autosave_debounce.setSingleShot(True)
        self._autosave_debounce.setInterval(1200)
        self._autosave_debounce.timeout.connect(self._save_autosave)
        self._autosave_periodic = QTimer(self)
        self._autosave_periodic.setInterval(15000)
        self._autosave_periodic.timeout.connect(self._save_autosave)

        # stats for practice
        self.practice_correct: Set[int] = set()
        self.practice_attempted: Set[int] = set()

        # wrong bank (persistent)
        self.wrong_bank: Set[int] = set(self._load_wrong_bank())

        # "last attempt incorrect" (ephemeral)
        self.last_incorrect: Set[int] = set()

        # AI Coach: per-scope queue of qids already analyzed (persistent)
        self.ai_coach_sent: Dict[str, Set[int]] = self._load_ai_coach_sent()
        self.ai_coach_batch_size: int = int(self.settings.value("ai_coach/batch_size", 10) or 10)
        self.ai_provider: str = self.settings.value("ai/provider", AI_PROVIDER_CLAUDE, str) or AI_PROVIDER_CLAUDE
        if self.ai_provider not in AI_MODELS_BY_PROVIDER:
            self.ai_provider = AI_PROVIDER_CLAUDE
        self.ai_model: str = (
            self.settings.value(f"ai/model/{self.ai_provider}", "", str)
            or self.settings.value("ai/model", "", str)
            or ""
        )
        self._ai_coach_thread: Optional[QThread] = None
        self._ai_coach_worker: Optional[AICoachWorker] = None
        self._ai_coach_progress: Optional[QProgressDialog] = None
        self._ai_coach_heartbeat: Optional[QTimer] = None
        self._ai_coach_state: Optional[Dict[str, object]] = None


        # Per-question stats: qid -> {'attempts': int, 'wrong': int, 'correct': int}
        self.q_stats: Dict[int, Dict[str, int]] = self._load_q_stats()
        self.p_stats: Dict[str, Dict[str, int]] = self._load_agg_stats("stats/per_pattern")
        self.t_stats: Dict[str, Dict[str, int]] = self._load_agg_stats("stats/per_tag")
        self._session_scored: Set[int] = set()  # legacy; kept for UI state
        self._last_scored_sel: Dict[int, FrozenSet[str]] = {}  # qid -> last selection that was counted as an attempt

        # background loading
        self._thread: Optional[QThread] = None
        self._worker: Optional[LoadWorker] = None

        self._build_ui()
        self._restore_last()

    # ---------- UI ----------
    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)

        top = QHBoxLayout()
        self.lbl_doc = QLabel("DOCX: (not loaded)")
        self.lbl_over = QLabel("Overrides: (none)")
        self.lbl_doc.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_over.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self.lbl_progress = QLabel("0 / 0")
        self.lbl_progress.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.lbl_timer = QLabel("\u23f1 --:--")
        self.lbl_timer.setFont(QFont("Segoe UI", 11, QFont.Bold))

        top.addWidget(self.lbl_doc, 1)
        top.addWidget(self.lbl_over, 1)
        top.addStretch(1)
        top.addWidget(self.lbl_progress)
        top.addSpacing(16)
        top.addWidget(self.lbl_timer)
        outer.addLayout(top)

        split = QSplitter(Qt.Horizontal)
        outer.addWidget(split, 1)

        # Left panel
        left = QWidget()
        left.setMinimumWidth(220)
        ll = QVBoxLayout(left)
        ll.setContentsMargins(6, 6, 6, 6)

        row = QHBoxLayout()
        self.btn_load = QPushButton("Load DOCX...")
        self.btn_load.clicked.connect(self.pick_docx)
        self.btn_over = QPushButton("Load Overrides...")
        self.btn_over.clicked.connect(self.pick_overrides)
        row.addWidget(self.btn_load)
        row.addWidget(self.btn_over)
        ll.addLayout(row)

        row_state = QHBoxLayout()
        row_state.addWidget(QLabel("Profile:"))
        self.cb_state_profile = QComboBox()
        self.cb_state_profile.setEditable(True)
        self.cb_state_profile.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.cb_state_profile.setToolTip("Type a profile name to save the current session, or pick one to restore it later.")
        row_state.addWidget(self.cb_state_profile, 1)

        self.btn_save_state = QPushButton("Save state")
        self.btn_save_state.clicked.connect(self.save_named_state)
        self.btn_load_state = QPushButton("Load state")
        self.btn_load_state.clicked.connect(self.load_named_state)
        self.btn_delete_state = QPushButton("Delete")
        self.btn_delete_state.clicked.connect(self.delete_named_state)
        row_state.addWidget(self.btn_save_state)
        row_state.addWidget(self.btn_load_state)
        row_state.addWidget(self.btn_delete_state)
        ll.addLayout(row_state)

        self.lbl_autosave = QLabel("Autosave: waiting for a session")
        self.lbl_autosave.setStyleSheet("opacity: 0.75;")
        ll.addWidget(self.lbl_autosave)

        row2 = QHBoxLayout()
        self.cb_mode = QComboBox()
        self.cb_mode.addItems(["Exam", "Practice"])
        self.cb_mode.currentTextChanged.connect(self._set_mode)

        self.sp_time = QSpinBox()
        self.sp_time.setRange(10, 400)
        self.sp_time.setValue(self.time_limit_min)
        if hasattr(self, 'sp_count'):
            self.sp_count.setValue(int(self.session_count))
        self.sp_time.valueChanged.connect(self._set_time_limit)

        row2.addWidget(QLabel("Mode:"))
        row2.addWidget(self.cb_mode, 1)
        row2.addSpacing(8)
        row2.addWidget(QLabel("Minutes:"))
        row2.addWidget(self.sp_time)
        ll.addLayout(row2)

        row3 = QHBoxLayout()
        self.cb_bank = QComboBox()
        self.cb_bank.addItems([
            "All questions",
            "New Qs + random mix",
            "New Qs only (306-322)",
            "Wrong >= Min",
            "Wrong bank (saved)",
            "Incorrect (last submit)",
            "Flagged",
            "Unanswered",
        ])
        self.cb_bank.setCurrentIndex(0)

        # Rebuild session when changing bank (so Wrong >= Min applies immediately)
        try:
            self.cb_bank.currentTextChanged.connect(lambda _=None: self.start_session() if self.questions else None)
        except Exception:
            pass

        self.chk_shuffle_q = QCheckBox("Shuffle Qs")
        self.chk_shuffle_q.setChecked(True)
        self.chk_shuffle_q.stateChanged.connect(lambda _=None: self._sync_shuffle_flags())

        self.chk_shuffle_o = QCheckBox("Shuffle options")
        self.chk_shuffle_o.setChecked(True)
        self.chk_shuffle_o.stateChanged.connect(lambda _=None: self._sync_shuffle_flags())

        row3.addWidget(self.cb_bank, 1)
        row3.addWidget(self.chk_shuffle_q)
        row3.addWidget(self.chk_shuffle_o)
        ll.addLayout(row3)

        row3c = QHBoxLayout()
        self.btn_new_mix = QPushButton("New Qs + Mix")
        self.btn_new_mix.setToolTip("Include questions 306-322, then fill the rest with random questions. Uses Count when Count > 0.")
        self.btn_new_mix.clicked.connect(self.start_new_mix_session)
        self.btn_new_only = QPushButton("New Qs Only")
        self.btn_new_only.setToolTip("Run only the new questions 306-322.")
        self.btn_new_only.clicked.connect(self.start_new_only_session)
        row3c.addWidget(self.btn_new_mix)
        row3c.addWidget(self.btn_new_only)
        ll.addLayout(row3c)

        # Wrong threshold + custom IDs (optional)
        row3b = QHBoxLayout()
        row3b.addWidget(QLabel("Min wrong >="))
        self.sp_min_wrong = QSpinBox()
        self.sp_min_wrong.setRange(0, 999)
        self.sp_min_wrong.setValue(int(self.min_wrong))
        self.sp_min_wrong.setToolTip("Filter for 'Wrong >= Min': include questions with wrong count >= this value. Use 0 for no minimum.")
        self.sp_min_wrong.valueChanged.connect(self._set_min_wrong)
        # Rebuild session when threshold changes (so the filter updates instantly)
        try:
            self.sp_min_wrong.valueChanged.connect(lambda _=None: self.start_session() if self.questions else None)
        except Exception:
            pass
        row3b.addWidget(self.sp_min_wrong)

        row3b.addSpacing(10)
        row3b.addWidget(QLabel("IDs:"))
        self.le_ids = QLineEdit()
        self.le_ids.setPlaceholderText("e.g., 1,2,4,297,305")
        self.le_ids.setText(self.custom_ids_text)
        self.le_ids.setToolTip("Comma-separated question IDs. If filled, the session will include ONLY these questions.")
        self.le_ids.editingFinished.connect(self._set_custom_ids_text)
        row3b.addWidget(self.le_ids, 1)

        btn_clear_ids = QToolButton()
        btn_clear_ids.setText("\u2716")
        btn_clear_ids.setToolTip("Clear IDs")
        btn_clear_ids.clicked.connect(lambda: self.le_ids.setText(""))
        row3b.addWidget(btn_clear_ids)
        ll.addLayout(row3b)

        row4 = QHBoxLayout()
        row4.addWidget(QLabel("Count:"))
        self.sp_count = QSpinBox()
        self.sp_count.setRange(0, 5000)
        self.sp_count.setValue(int(self.session_count))
        self.sp_count.setToolTip("How many questions to include in this session (0 = ALL).")
        self.sp_count.valueChanged.connect(self._set_session_count)
        row4.addWidget(self.sp_count)
        row4.addStretch(1)
        row4.addWidget(QLabel("0 = ALL"))
        ll.addLayout(row4)

        self.list_q = QListWidget()
        self.list_q.itemClicked.connect(self._jump_to_clicked)
        ll.addWidget(self.list_q, 1)

        # left buttons
        self.btn_start = QPushButton("Start / Restart")
        self.btn_start.clicked.connect(self.start_session)

        self.btn_submit = QPushButton("Submit")
        self.btn_submit.clicked.connect(self.submit_or_finish)

        self.btn_clear_bank = QPushButton("Clear wrong bank")
        self.btn_clear_bank.clicked.connect(self.clear_wrong_bank)

        self.btn_export = QPushButton("Export Overrides Template...")
        self.btn_export.clicked.connect(self.export_template)

        ll.addWidget(self.btn_start)

        row_wrong = QHBoxLayout()
        self.btn_do_wrong = QPushButton("DB wrong answers")
        self.btn_do_wrong.setToolTip("Drill the wrong-answer DB: saved bank + AI Coach + Deep Review qids.")
        self.btn_do_wrong.clicked.connect(self.start_wrong_answers)
        self.btn_round_wrong = QPushButton("Round wrong answers")
        self.btn_round_wrong.setToolTip("Drill only the questions you got wrong in the CURRENT round (live selections).")
        self.btn_round_wrong.clicked.connect(self.start_round_wrong_answers)
        self.btn_repeat_all = QPushButton("Repeat all questions")
        self.btn_repeat_all.setToolTip("Restart the current loaded/profile question set, regardless of wrong or correct status.")
        self.btn_repeat_all.clicked.connect(self.start_repeat_all_questions)
        self.btn_stats = QPushButton("Stats...")
        self.btn_stats.clicked.connect(self.show_stats)
        self.btn_ai_coach = QPushButton("\U0001f9e0 AI Coach...")
        self.btn_ai_coach.setToolTip("Ask the selected AI provider to diagnose why you keep missing these questions.")
        self.btn_ai_coach.clicked.connect(self.show_ai_coach)
        self.btn_meta_coach = QPushButton("\U0001f5fa\ufe0f Meta-Coach")
        self.btn_meta_coach.setToolTip("Read the entire ai_coach_reports.md history and produce a meta-analysis + study plan.")
        self.btn_meta_coach.clicked.connect(self.show_meta_coach)
        self.btn_db_report = QPushButton("\U0001f4da Show AI Review...")
        self.btn_db_report.setToolTip("Abre el ai_coach_reports.md entero (consolidado, sin Claude call).")
        self.btn_db_report.clicked.connect(self.show_db_report)
        self.btn_show_deep = QToolButton()
        self.btn_show_deep.setText("\U0001f4d6 Show Deep Review...")
        self.btn_show_deep.setToolTip("Abre el deep_review_reports.md entero (sin Claude call). Menu para filtrar al row actual.")
        self.btn_show_deep.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        self.btn_show_deep.clicked.connect(self.show_deep_review_md)
        _menu_deep = QMenu(self.btn_show_deep)
        _menu_deep.addAction("Show all (default)", self.show_deep_review_md)
        _menu_deep.addAction("Show current round (all qids)", self.show_deep_review_md_current_round)
        _menu_deep.addAction("Show only current row wrongs", self.show_deep_review_md_current_row)
        self.btn_show_deep.setMenu(_menu_deep)
        self.btn_deep_review = QPushButton("\U0001f52c Deep Review...")
        self.btn_deep_review.setToolTip("C\u00e1tedra completa por pregunta. Caro: ~2-3 min/qid.")
        self.btn_deep_review.clicked.connect(self.show_deep_review)
        self.btn_pre_brief = QPushButton("\U0001f4cb Pre-Brief...")
        self.btn_pre_brief.setToolTip("Dossier de contexto barato por qid. Prepara terreno antes de Deep Review.")
        self.btn_pre_brief.clicked.connect(self.show_pre_brief)
        self.btn_dossiers = QPushButton("\U0001f5c2\ufe0f Build Dossiers...")
        self.btn_dossiers.setToolTip("Consolida deep_review_reports.md en dossiers canonicos por concepto AWS.")
        self.btn_dossiers.clicked.connect(self.build_concept_dossiers)
        self.btn_show_dossiers = QToolButton()
        self.btn_show_dossiers.setText("\U0001f4da Dossiers...")
        self.btn_show_dossiers.setToolTip("Abre un dossier de concepto guardado en concept_dossiers/ (sin Claude call).")
        self.btn_show_dossiers.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self._dossier_menu = QMenu(self.btn_show_dossiers)
        self._dossier_menu.aboutToShow.connect(self._rebuild_concept_dossier_menu)
        self.btn_show_dossiers.setMenu(self._dossier_menu)
        self.btn_rag_kb = QPushButton("RAG KB")
        self.btn_rag_kb.setToolTip(
            "Busca contexto en la knowledge base local, arma un prompt grounded con citas "
            "y lo pasa al provider seguro (dry-run unless AI_LIVE=1)."
        )
        self.btn_rag_kb.clicked.connect(self.show_rag_kb)
        self.btn_nuclear = QPushButton("\u2622 Nuclear")
        self.btn_nuclear.setToolTip("Ultradios mode: 3 agentes paralelos + sintetizador para UNA pregunta critica.")
        self.btn_nuclear.clicked.connect(self.show_nuclear_review)
        self.btn_show_nuclear = QPushButton("\u2622 Show Nuclear...")
        self.btn_show_nuclear.setToolTip("Abre el nuclear_reports.md entero (sin Claude call).")
        self.btn_show_nuclear.clicked.connect(self.show_nuclear_md)
        self.btn_diagram = QPushButton("\U0001f4ca Diagram")
        self.btn_diagram.setToolTip(
            "Genera un HTML de estudio (multi-diagramas + sequence diagrams) para la pregunta actual.\n"
            "Lee artifacts previos (deep / nuclear / ai_coach / Q<N>_*) si existen para no re-derivar de cero.\n"
            "Guarda en Q<N>_diagram.html y abre en el browser."
        )
        self.btn_diagram.clicked.connect(self.show_diagram)
        self.btn_teach_zero = QPushButton("\U0001f393 C\u00e1tedra")
        self.btn_teach_zero.setToolTip(
            "Tier 5 - Lecture markdown 'Teach Me From Zero' para la pregunta actual.\n"
            "Lee artifacts previos (deep / nuclear / ai_coach / Q<N>_*) si existen.\n"
            "Salida en popup con markdown renderizado y guarda Q<N>_teach_<ts>.md."
        )
        self.btn_teach_zero.clicked.connect(self.show_teach_zero)
        self.btn_top_offenders = QPushButton("\U0001f525 Top Offenders")
        self.btn_top_offenders.setToolTip("Lista de las preguntas que mas has fallado (lifetime wrong count). Click en Q### para saltar.")
        self.btn_top_offenders.clicked.connect(self.show_top_offenders)
        self.btn_due = QPushButton("\U0001f5d3\ufe0f Due Today")
        self.btn_due.setToolTip("Drill SRS - preguntas cuya fecha next_review llego (srs_state.json).")
        self.btn_due.clicked.connect(self.start_due_today)
        self.btn_mastery = QPushButton("\U0001f3af Mastery")
        self.btn_mastery.setToolTip("Mastery por concepto AWS (concept_mastery.json). Decide donde gastar Nuclear.")
        self.btn_mastery.clicked.connect(self.show_concept_mastery)
        row_wrong.addWidget(self.btn_do_wrong)
        row_wrong.addWidget(self.btn_round_wrong)
        row_wrong.addWidget(self.btn_repeat_all)
        row_wrong.addWidget(self.btn_stats)
        row_wrong.addWidget(self.btn_top_offenders)
        row_wrong.addWidget(self.btn_due)
        row_wrong.addWidget(self.btn_mastery)
        ll.addLayout(row_wrong)

        row_ai_backend = QHBoxLayout()
        row_ai_backend.addWidget(QLabel("AI:"))
        self.cb_ai_provider = QComboBox()
        self.cb_ai_provider.addItems([AI_PROVIDER_LABELS[AI_PROVIDER_CLAUDE], AI_PROVIDER_LABELS[AI_PROVIDER_OPENAI]])
        self.cb_ai_provider.setToolTip("Provider usado por AI Coach, Deep, Nuclear, Diagram, C\u00e1tedra y Dossiers.")
        self.cb_ai_provider.setCurrentText(AI_PROVIDER_LABELS.get(self.ai_provider, AI_PROVIDER_LABELS[AI_PROVIDER_CLAUDE]))
        self.cb_ai_provider.currentTextChanged.connect(self._set_ai_provider)
        row_ai_backend.addWidget(self.cb_ai_provider)

        self.cb_ai_model = QComboBox()
        self.cb_ai_model.setEditable(True)
        self.cb_ai_model.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.cb_ai_model.setToolTip("Modelo para el proveedor seleccionado. OpenAI usa OPENAI_API_KEY o un OPENAI_BASE_URL local.")
        row_ai_backend.addWidget(self.cb_ai_model, 1)
        self._refresh_ai_model_choices(self.ai_model)
        self.cb_ai_model.currentTextChanged.connect(self._set_ai_model)
        ll.addLayout(row_ai_backend)

        row_rag = QHBoxLayout()
        row_rag.addWidget(QLabel("Knowledge Base:"))
        self.btn_rag_kb.setText("RAG KB / Grounded Context")
        self.btn_rag_kb.setMinimumHeight(32)
        self.btn_rag_kb.setStyleSheet("font-weight: 600;")
        row_rag.addWidget(self.btn_rag_kb)
        row_rag.addStretch(1)
        ll.addLayout(row_rag)

        row_coach = QHBoxLayout()
        row_coach.addWidget(self.btn_ai_coach)
        row_coach.addWidget(self.btn_meta_coach)
        row_coach.addWidget(self.btn_pre_brief)
        ll.addLayout(row_coach)

        row_heavy = QHBoxLayout()
        row_heavy.addWidget(self.btn_deep_review)
        row_heavy.addWidget(self.btn_show_deep)
        row_heavy.addWidget(self.btn_db_report)
        row_heavy.addWidget(self.btn_dossiers)
        row_heavy.addWidget(self.btn_show_dossiers)
        row_heavy.addWidget(self.btn_nuclear)
        row_heavy.addWidget(self.btn_show_nuclear)
        row_heavy.addWidget(self.btn_diagram)
        row_heavy.addWidget(self.btn_teach_zero)
        ll.addLayout(row_heavy)

        ll.addWidget(self.btn_submit)
        ll.addWidget(self.btn_clear_bank)
        ll.addWidget(self.btn_export)

        split.addWidget(left)
        split.setStretchFactor(0, 0)

        # Right panel
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(10, 10, 10, 10)

        hdr = QHBoxLayout()
        self.lbl_qtitle = QLabel("Question")
        self.lbl_qtitle.setFont(QFont("Segoe UI", 12, QFont.Bold))
        hdr.addWidget(self.lbl_qtitle, 1)

        self.btn_flag = QToolButton()
        self.btn_flag.setText("\u2691")
        self.btn_flag.clicked.connect(self.toggle_flag)

        self.btn_clear = QToolButton()
        self.btn_clear.setText("Clear")
        self.btn_clear.clicked.connect(self.clear_selection)

        self.cb_confidence = QComboBox()
        self.cb_confidence.addItems(["Sure", "Doubt", "Guess"])
        self.cb_confidence.setToolTip(
            "Tu confianza al responder. Se loguea por intento en confidence_log.jsonl.\n"
            "Sirve para detectar lucky guesses (right+Guess) y misconceptions (wrong+Sure)."
        )
        self.cb_confidence.setCurrentIndex(0)

        hdr.addWidget(self.cb_confidence)
        hdr.addWidget(self.btn_clear)
        hdr.addWidget(self.btn_flag)
        rl.addLayout(hdr)

        self.txt_stem = QTextBrowser()
        self.txt_stem.setOpenExternalLinks(False)
        # Keep the question box compact so the options area gets most of the space.
        # (Prevents unnecessary scrolling in the answers when the question is short.)
        self.txt_stem.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        self.txt_stem.setMinimumHeight(120)
        self.txt_stem.setMaximumHeight(280)
        rl.addWidget(self.txt_stem, 0)

        self.options_area = QScrollArea()
        self.options_area.setWidgetResizable(True)
        self.options_host = QWidget()
        self.options_l = QVBoxLayout(self.options_host)
        self.options_l.setContentsMargins(0, 0, 0, 0)
        self.options_l.setSpacing(8)
        self.options_l.addStretch(1)
        self.options_area.setWidget(self.options_host)
        self.options_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        rl.addWidget(self.options_area, 3)

        # Feedback / explanation
        self.txt_feedback = QTextBrowser()

        self.txt_feedback.setVisible(False)

        # Make explanation readable without horizontal scrolling
        try:
            self.txt_feedback.setLineWrapMode(QTextEdit.WidgetWidth)
        except Exception:
            pass
        self.txt_feedback.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.txt_feedback.setMinimumHeight(220)
        rl.addWidget(self.txt_feedback, 2)
# nav row
        nav = QHBoxLayout()
        self.btn_prev = QPushButton("\u25c0 Prev")
        self.btn_prev.clicked.connect(self.prev_q)
        self.btn_next = QPushButton("Next \u25b6")
        self.btn_next.clicked.connect(self.next_q)

        self.chk_auto_next = QCheckBox("Auto-check on Next")
        self.chk_auto_next.setChecked(bool(self.auto_check_next_enabled))
        self.chk_auto_next.setToolTip("Practice only: when enabled, Next will auto-check (then Next again advances). Wrong answers get auto-flagged for review.")
        self.chk_auto_next.stateChanged.connect(self._set_auto_check_next)

        self.btn_check = QPushButton("Check (Practice)")
        self.btn_check.clicked.connect(self.check_practice)

        nav.addWidget(self.btn_prev)
        nav.addWidget(self.btn_next)
        nav.addStretch(1)
        nav.addWidget(self.chk_auto_next)
        nav.addWidget(self.btn_check)
        rl.addLayout(nav)

        split.addWidget(right)
        split.setStretchFactor(1, 1)

        # Give the right (question/answers) panel more space by default.
        try:
            split.setSizes([360, 1000])
            split.setCollapsible(0, False)
            split.setCollapsible(1, False)
            split.setHandleWidth(10)
            split.setOpaqueResize(True)
        except Exception:
            pass

        self.lbl_hint = QLabel("Load your DOCX to begin.")
        self.lbl_hint.setWordWrap(True)
        self.lbl_hint.setStyleSheet("opacity: 0.85;")
        outer.addWidget(self.lbl_hint)

        self._refresh_profile_combo()
        self._set_mode(self.mode)

    def _refresh_ai_model_choices(self, preferred: str = "") -> None:
        if not hasattr(self, "cb_ai_model"):
            return
        provider = getattr(self, "ai_provider", AI_PROVIDER_CLAUDE)
        models = list(AI_MODELS_BY_PROVIDER.get(provider, AI_MODELS_BY_PROVIDER[AI_PROVIDER_CLAUDE]))
        preferred = (preferred or self.ai_model or "").strip()
        if preferred and preferred not in models:
            models.insert(0, preferred)

        old_block = self.cb_ai_model.blockSignals(True)
        try:
            self.cb_ai_model.clear()
            self.cb_ai_model.addItems(models)
            chosen = preferred if preferred in models else (models[0] if models else "")
            if chosen:
                self.cb_ai_model.setCurrentText(chosen)
                self.ai_model = chosen
        finally:
            self.cb_ai_model.blockSignals(old_block)
        self.settings.setValue("ai/model", self.ai_model)
        self.settings.setValue(f"ai/model/{provider}", self.ai_model)

    def _set_ai_provider(self, label: str) -> None:
        provider = AI_PROVIDER_BY_LABEL.get(label, AI_PROVIDER_CLAUDE)
        self.ai_provider = provider
        self.settings.setValue("ai/provider", provider)
        saved_for_provider = self.settings.value(f"ai/model/{provider}", "", str) or ""
        self.ai_model = saved_for_provider
        self._refresh_ai_model_choices(saved_for_provider)
        self.settings.sync()

    def _set_ai_model(self, model: str) -> None:
        self.ai_model = (model or "").strip()
        self.settings.setValue("ai/model", self.ai_model)
        self.settings.setValue(f"ai/model/{self.ai_provider}", self.ai_model)
        self.settings.sync()

    def _selected_ai_backend(self) -> Tuple[str, Optional[str]]:
        provider = getattr(self, "ai_provider", AI_PROVIDER_CLAUDE)
        if hasattr(self, "cb_ai_provider"):
            provider = AI_PROVIDER_BY_LABEL.get(self.cb_ai_provider.currentText(), provider)
        model = getattr(self, "ai_model", "") or ""
        if hasattr(self, "cb_ai_model"):
            model = self.cb_ai_model.currentText().strip()
        send_model = model
        if provider == AI_PROVIDER_CLAUDE and model.startswith("("):
            send_model = ""
        self.ai_provider = provider
        self.ai_model = model
        return provider, (send_model or None)

    def _ai_backend_label(self) -> str:
        provider, model = self._selected_ai_backend()
        return f"{AI_PROVIDER_LABELS.get(provider, provider)} - {model or 'default'}"

    def _provider_report_path(self, kind: str, provider: Optional[str] = None) -> Path:
        p = provider or getattr(self, "ai_provider", AI_PROVIDER_CLAUDE)
        base_name, openai_name = AI_REPORT_FILES[kind]
        return Path(__file__).with_name(openai_name if p == AI_PROVIDER_OPENAI else base_name)

    def _provider_dossier_dir(self, provider: Optional[str] = None) -> Path:
        p = provider or getattr(self, "ai_provider", AI_PROVIDER_CLAUDE)
        name = "concept_dossiers_openai" if p == AI_PROVIDER_OPENAI else "concept_dossiers"
        return Path(__file__).with_name(name)

    def _provider_artifact_token(self, provider: Optional[str] = None) -> str:
        p = provider or getattr(self, "ai_provider", AI_PROVIDER_CLAUDE)
        return "_openai" if p == AI_PROVIDER_OPENAI else ""

    def _ensure_ai_backend_ready(self) -> bool:
        provider, _model = self._selected_ai_backend()
        base_url = (os.environ.get("OPENAI_BASE_URL") or "").strip()
        local_gateway = bool(re.match(r"^https?://(?:127\.0\.0\.1|localhost|\[::1\])(?::\d+)?(?:/|$)", base_url, re.I))
        if provider == AI_PROVIDER_OPENAI and not local_gateway and not (os.environ.get("OPENAI_API_KEY") or "").strip():
            from PySide6.QtWidgets import QInputDialog
            key, ok = QInputDialog.getText(
                self,
                "OpenAI API key missing",
                "OpenAI is selected but this app cannot see OPENAI_API_KEY.\n\n"
                "Paste the key for this app session only. It will not be saved.",
                QLineEdit.EchoMode.Password,
            )
            key = (key or "").strip()
            if not ok or not key:
                QMessageBox.information(
                    self,
                    "OpenAI API key missing",
                    "No key loaded. You can also launch the app from PowerShell:\n\n"
                    "$env:OPENAI_API_KEY = \"sk-...\"\n"
                    "python ans_c01_quiz_gui_v2_counterfix2.py\n\n"
                    "The key is not stored in code or QSettings.",
                )
                return False
            os.environ["OPENAI_API_KEY"] = key
        return True

    def _ai_error_for_dialog(self, error: str) -> str:
        text = str(error or "")
        lower = text.lower()
        if "insufficient_quota" in lower or "exceeded your current quota" in lower:
            return (
                "OpenAI rejected the request with HTTP 429: insufficient_quota.\n\n"
                "This is a billing/quota problem on the API project, not a bug in the app. "
                "Add billing/credits to the OpenAI project, use a different API key/project, "
                "or switch the AI dropdown back to Claude.\n\n"
                "No OpenAI report was saved for this failed call."
            )
        if "openai api http 429" in lower or "too many requests" in lower:
            return (
                "OpenAI returned HTTP 429.\n\n"
                "If the error code is insufficient_quota, add billing/credits or use another key. "
                "If it is a rate limit, wait a bit and retry with a smaller batch."
            )
        return text[:2000]

    def _sync_shuffle_flags(self):
        self.shuffle_questions = self.chk_shuffle_q.isChecked()
        self.shuffle_options = self.chk_shuffle_o.isChecked()

    # ---------- persistence ----------
    def _load_wrong_bank(self) -> List[int]:
        raw = self.settings.value("bank/wrong", "[]", str) or "[]"
        try:
            data = json.loads(raw)
            if isinstance(data, list):
                return [int(x) for x in data if str(x).isdigit()]
        except Exception:
            pass
        return []

    def _save_wrong_bank(self) -> None:
        self.settings.setValue("bank/wrong", json.dumps(sorted(self.wrong_bank)))
        self.settings.sync()

    def _wrong_db_qids(self) -> Set[int]:
        """Canonical wrong-answer DB used by drills and coach scopes.

        The old saved bank can be tiny if QSettings was reset or a profile changed,
        but AI Coach / Deep Review archives are also durable evidence that a qid
        belonged to the wrong-answer workflow. Use the union so "DB wrong answers"
        matches the report DB the user sees.

        Excludes qids the user has now mastered: current correct streak >= 3.
        Streak resets to 0 on any wrong answer, so a single miss puts the qid
        back in the wrong DB and AI Coach / Deep Review can re-engage.
        """
        raw = set(self.wrong_bank) | self._load_master_qids() | self._load_deep_qids()
        mastered = {
            qid for qid, s in self.q_stats.items()
            if int(s.get("streak", 0) or 0) >= 3
        }
        return raw - mastered

    def _load_ai_coach_sent(self) -> Dict[str, Set[int]]:
        raw = self.settings.value("ai_coach/sent", "{}", str) or "{}"
        try:
            data = json.loads(raw)
            if isinstance(data, dict):
                return {
                    str(k): {int(x) for x in v if str(x).lstrip("-").isdigit()}
                    for k, v in data.items() if isinstance(v, list)
                }
        except Exception:
            pass
        return {}

    def _save_ai_coach_sent(self) -> None:
        payload = {k: sorted(v) for k, v in self.ai_coach_sent.items()}
        self.settings.setValue("ai_coach/sent", json.dumps(payload))
        self.settings.sync()

    def _load_master_qids(self) -> Set[int]:
        """Parse ai_coach_reports.md and return every qid actually analyzed.

        Source of truth for dedup - survives QSettings resets and scope changes.
        We extract qids from the `### Q<N>` section headings, NOT from the
        `<!-- AI_COACH_BATCH qids=[...] -->` marker, because the marker reflects
        what was REQUESTED while the headings reflect what Claude actually emitted
        (model can drop qids on output limits or timeouts).
        """
        master_path = self._provider_report_path("coach")
        if not master_path.exists():
            return set()
        try:
            text = master_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return set()
        out: Set[int] = set()
        for m in re.finditer(r"^###\s+Q(\d+)\b", text, re.MULTILINE):
            out.add(int(m.group(1)))
        return out

    def _load_master_qid_answers(self) -> Dict[int, Set[str]]:
        """Per-qid set of wrong-answer letters previously analyzed.

        Parses each `### Q<id>` section's `**Your answer:**` line in
        ai_coach_reports.md. Used to detect when the user has flipped to a
        different wrong option for a qid that was already analyzed - that's a
        'no stable model' signal worth re-engaging Claude on.
        """
        master_path = self._provider_report_path("coach")
        if not master_path.exists():
            return {}
        try:
            text = master_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return {}
        out: Dict[int, Set[str]] = {}
        for m in re.finditer(
            r"###\s*Q(\d+)\b[^\n]*\n(?:[^\n]*\n){0,3}?[^\n]*\*\*Your answer:\*\*\s*([A-Za-z](?:\s*,\s*[A-Za-z])*)",
            text,
        ):
            try:
                qid = int(m.group(1))
            except ValueError:
                continue
            ans = re.sub(r"\s+", "", m.group(2)).upper()
            out.setdefault(qid, set()).add(ans)
        return out

    def _inject_recaps_inline(self, text: str) -> str:
        """For each `### Q<id>` section without a recap, inject one built from
        the loaded docx (stem + options inline). Pure local rewrite - no tokens.

        Sections whose qid is not in the currently loaded bank get a small note
        instead, since we have no source for the stem.
        """
        if not text:
            return text
        loaded: Dict[int, "Question"] = {q.qid: q for q in (self.questions or [])}

        def _short(s: str, n: int) -> str:
            s = (s or "").strip().replace("\r", " ").replace("\n", " ")
            s = re.sub(r"\s+", " ", s)
            return s if len(s) <= n else s[: n - 1].rstrip() + "..."

        def _recap_for(qid: int) -> Optional[str]:
            q = loaded.get(qid)
            if q is None:
                return None
            stem_short = _short(q.stem or "", 600)
            opts = q.options or {}
            opts_inline = " \u00b7 ".join(
                f"{letter}) {_short(opts[letter], 120)}"
                for letter in sorted(opts.keys())
            )
            return f"**Question recap:** {stem_short}\n{opts_inline}\n\n"

        out_parts: List[str] = []
        sections = re.split(r"(?=^###\s+Q\d+)", text, flags=re.MULTILINE)
        for sec in sections:
            m = re.match(r"^###\s+Q(\d+)\b[^\n]*\n", sec)
            if not m:
                out_parts.append(sec)
                continue
            if "**Question recap:**" in sec or "**Question:**" in sec:
                out_parts.append(sec)
                continue
            try:
                qid = int(m.group(1))
            except ValueError:
                out_parts.append(sec)
                continue
            recap = _recap_for(qid)
            if recap is None:
                note = f"_(recap unavailable: Q{qid} not in currently loaded docx)_\n\n"
                out_parts.append(sec[: m.end()] + note + sec[m.end():])
            else:
                out_parts.append(sec[: m.end()] + recap + sec[m.end():])
        return "".join(out_parts)

    def _load_deep_qids(self) -> Set[int]:
        """Parse deep_review_reports.md and return every qid actually deep-reviewed.

        Source of truth: per-qid `<!-- DEEP_REVIEW_QID qid=N ... -->` markers,
        which are written only on successful per-qid completion. We deliberately
        ignore the batch-level `DEEP_REVIEW qids=[...]` marker because it
        reflects what was REQUESTED, not what landed on disk (Claude can drop
        qids on output limits).
        """
        deep_path = self._provider_report_path("deep")
        if not deep_path.exists():
            return set()
        try:
            text = deep_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return set()
        out: Set[int] = set()
        for m in re.finditer(r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)", text):
            out.add(int(m.group(1)))
        return out

    @Slot(str)
    def _ai_coach_progress_update(self, text: str) -> None:
        if not self._ai_coach_state:
            return
        if bool(self._ai_coach_state.get("cancelled", False)):
            return
        self._ai_coach_state["last_msg"] = text

    @Slot()
    def _tick_ai_coach_heartbeat(self) -> None:
        if not self._ai_coach_state or not self._ai_coach_progress:
            return
        tick = int(self._ai_coach_state.get("tick", 0) or 0) + 1
        self._ai_coach_state["tick"] = tick
        # Pulse the busy bar (forces Windows to keep the marquee animated)
        try:
            self._ai_coach_progress.setValue(0)
        except Exception:
            pass
        # Only re-render the label when the underlying text actually changes,
        # to avoid hammering WM_PAINT and triggering "Not Responding".
        last_msg = str(self._ai_coach_state.get("last_msg", "waiting for first chunk..."))
        scope_label = str(self._ai_coach_state.get("scope_label", ""))
        # Deep Review parallel mode: show per-qid checklist
        qstat = self._ai_coach_state.get("qid_status") or {}
        status_line = ""
        if qstat:
            done = sum(1 for v in qstat.values() if v == "done")
            err = sum(1 for v in qstat.values() if v == "error")
            pend = [q for q, v in sorted(qstat.items()) if v == "pending"]
            done_qs = [q for q, v in sorted(qstat.items()) if v == "done"]
            err_qs = [q for q, v in sorted(qstat.items()) if v == "error"]
            parts = []
            if done_qs:
                parts.append("\u2705 " + " ".join(f"Q{q}" for q in done_qs))
            if err_qs:
                parts.append("\u274c " + " ".join(f"Q{q}" for q in err_qs))
            if pend:
                parts.append("\u23f3 " + " ".join(f"Q{q}" for q in pend))
            status_line = f"\n[{done}/{len(qstat)} done"
            if err:
                status_line += f", {err} err"
            status_line += "]  " + " \u00b7 ".join(parts)
        progress_title = "RAG KB" if bool(self._ai_coach_state.get("is_rag_kb", False)) else "AI Coach"
        new_label = f"{progress_title} - {tick}s elapsed\n{last_msg}\n({scope_label}){status_line}"
        prev_label = self._ai_coach_state.get("rendered_label", "")
        # Update at most every 3s, OR whenever the streaming message changed
        last_render_tick = int(self._ai_coach_state.get("last_render_tick", 0) or 0)
        if last_msg != self._ai_coach_state.get("rendered_msg", "") or (tick - last_render_tick) >= 3:
            self._ai_coach_progress.setLabelText(new_label)
            self._ai_coach_state["rendered_label"] = new_label
            self._ai_coach_state["rendered_msg"] = last_msg
            self._ai_coach_state["last_render_tick"] = tick

    @Slot()
    def _cancel_ai_coach(self) -> None:
        if not self._ai_coach_state or not self._ai_coach_progress:
            return
        self._ai_coach_state["cancelled"] = True
        if self._ai_coach_heartbeat:
            self._ai_coach_heartbeat.stop()
        self._ai_coach_progress.setLabelText("Cancelling... (waiting for in-flight call to return)")

    @Slot(int, str)
    def _deep_review_qid_done(self, qid: int, md: str) -> None:
        """Persist a single Deep Review qid result IMMEDIATELY as it arrives.

        This is the safety net: even if the worker later crashes, times out, or
        the user closes the app, every qid that finished is already on disk.
        """
        try:
            state = self._ai_coach_state or {}
            scope = str(state.get("scope", ""))
            already = set(state.get("incremental_saved", set()) or set())
            failed = bool(re.search(r"_Error:\s", (md or "")[:200]))
            if int(qid) in already:
                return
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            section = (
                f"\n\n<!-- DEEP_REVIEW_QID qid={int(qid)} scope={scope} ts={ts} status={'error' if failed else 'ok'} -->\n"
                f"# {ts} - Deep Review - Q{int(qid)} - scope: {scope}\n\n"
                f"{md}\n"
            )
            deep_path = self._provider_report_path("deep", str(state.get("ai_provider") or AI_PROVIDER_CLAUDE))
            with deep_path.open("a", encoding="utf-8") as f:
                f.write(section)
            already.add(int(qid))
            if isinstance(self._ai_coach_state, dict):
                self._ai_coach_state["incremental_saved"] = already
                qstat = self._ai_coach_state.get("qid_status") or {}
                qstat[int(qid)] = "error" if failed else "done"
                self._ai_coach_state["qid_status"] = qstat
                # Force heartbeat to re-render on next tick
                status_icon = "\u274c" if failed else "\u2705"
                self._ai_coach_state["last_msg"] = (
                    f"{status_icon} Q{int(qid)} saved "
                    f"({sum(1 for v in qstat.values() if v != 'pending')}/{len(qstat)} done)"
                )
        except Exception:
            import logging
            logging.getLogger("quiz_ai_coach.deep_worker").exception(
                "incremental save failed for Q%s", qid,
            )

    @Slot(str, str)
    def _ai_coach_finished(self, report_md: str, error: str) -> None:
        state = dict(self._ai_coach_state or {})
        progress = self._ai_coach_progress
        heartbeat = self._ai_coach_heartbeat
        worker = self._ai_coach_worker
        thread = self._ai_coach_thread

        self._ai_coach_progress = None
        self._ai_coach_heartbeat = None
        self._ai_coach_worker = None
        self._ai_coach_thread = None
        self._ai_coach_state = None

        if heartbeat:
            heartbeat.stop()
            heartbeat.deleteLater()
        if progress:
            progress.close()
            progress.deleteLater()
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        is_meta = bool(state.get("is_meta", False))
        is_deep = bool(state.get("is_deep", False))
        is_prebrief = bool(state.get("is_prebrief", False))
        try:
            if is_meta:
                self.btn_meta_coach.setEnabled(True)
            elif is_deep:
                self.btn_deep_review.setEnabled(True)
            elif is_prebrief and hasattr(self, "btn_pre_brief"):
                self.btn_pre_brief.setEnabled(True)
            else:
                self.btn_ai_coach.setEnabled(True)
        except Exception:
            pass

        if bool(state.get("cancelled", False)):
            return

        if error:
            import sys as _sys
            print(f"[AI Coach error]\n{error}", file=_sys.stderr)
            title_err = (
                "Meta-Coach failed" if is_meta else
                "Deep Review failed" if is_deep else
                "Pre-Brief failed" if is_prebrief else
                "AI Coach failed"
            )
            QMessageBox.critical(self, title_err, f"{self._ai_error_for_dialog(error)}\n\nSee ai_coach.log for details.")
            return
        if not report_md.strip():
            QMessageBox.warning(
                self, "AI Coach",
                f"Empty response from {state.get('ai_backend_label', 'AI')}. See ai_coach.log.",
            )
            return

        if is_meta:
            provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
            # Compute progress vs previous Meta-Coach run BEFORE saving the new
            # snapshot, so the diff compares against the truly previous one.
            stats_payload = state.get("meta_stats_payload") or {}
            progress_md = self._meta_progress_section(stats_payload, provider_for_save) if stats_payload else ""
            if stats_payload:
                self._append_meta_history(stats_payload, provider_for_save)

            full_report = progress_md + report_md
            # Meta-coach: save to its own file (overwrite each run; it's a snapshot)
            meta_path = self._provider_report_path("meta", provider_for_save)
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                meta_path.write_text(
                    f"# Meta-Coach analysis - {ts}\n\n{full_report}\n",
                    encoding="utf-8",
                )
                appended_msg = f"Saved to: {meta_path.name}"
                # Persist source hash so next click can short-circuit if reports.md
                # hasn't changed (saves a big-context Claude call).
                try:
                    import hashlib as _hashlib
                    src_path = self._provider_report_path("coach", provider_for_save)
                    if src_path.exists():
                        src_text = src_path.read_text(encoding="utf-8", errors="replace")
                        h = _hashlib.sha256(src_text.encode("utf-8")).hexdigest()
                        self._provider_report_path("meta_hash", provider_for_save).write_text(
                            h, encoding="utf-8",
                        )
                except Exception:
                    pass
            except Exception as _ex:
                appended_msg = f"(save failed: {_ex})"
            footer = f"\n\n---\n_{appended_msg}_"
            unique_q = int(state.get("meta_unique_qids", 0) or 0)
            batch_n = int(state.get("meta_batch_count", 0) or 0)
            pending_n = int(state.get("meta_pending", 0) or 0)
            hot_n = int(state.get("meta_hot_zone", 0) or 0)
            header_txt = (
                f"Meta-analysis over {unique_q} unique qids across {batch_n} batches "
                f"\u00b7 {hot_n} in hot zone \u00b7 {pending_n} pending in bank"
            )
            dlg = AICoachDialog(
                self, full_report + footer, num_questions=unique_q,
                model=f"{state.get('ai_backend_label', 'AI')} \u00b7 meta-coach",
                master_path=meta_path,
                header_text=header_txt,
            )
            dlg.setWindowTitle("Meta-Coach - Brain retraining plan")
            self._show_coach_dialog_nonmodal(dlg)
            return

        if is_deep:
            scope = str(state.get("scope", ""))
            scope_label = str(state.get("scope_label", ""))
            sent_n = int(state.get("sent_n", 0) or 0)
            sent_qids = sorted(state.get("sent_qids", set()))
            with_prior = int(state.get("with_prior", 0) or 0)

            provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
            deep_path = self._provider_report_path("deep", provider_for_save)
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            qids_str = ", ".join(str(q) for q in sent_qids)
            already_saved = set(state.get("incremental_saved", set()) or set())
            missing = [q for q in sent_qids if int(q) not in already_saved]
            try:
                # Each qid was already saved incrementally via _deep_review_qid_done.
                # We only append a small batch-summary marker so the existing dedup
                # (which reads `DEEP_REVIEW qids=[...]`) and Meta-Coach see the batch.
                summary = (
                    f"\n\n<!-- DEEP_REVIEW qids=[{qids_str}] scope={scope} ts={ts} "
                    f"incremental_saved={len(already_saved)}/{sent_n} -->\n"
                    f"_Deep Review batch summary \u00b7 scope **{scope}** \u00b7 {sent_n} qids "
                    f"(incrementally saved per-qid above)._\n"
                )
                if missing:
                    summary += f"\n_Note: {len(missing)} qids did NOT save incrementally (qids: {missing}); content below is the post-hoc fallback._\n\n{report_md}\n"
                with deep_path.open("a", encoding="utf-8") as f:
                    f.write(summary)
                appended_msg = f"Appended to: {deep_path.name} ({len(already_saved)}/{sent_n} saved incrementally)"
            except Exception as _ex:
                appended_msg = f"(append failed: {_ex})"
            footer = (
                f"\n\n---\n_Deep Review batch saved \u00b7 scope **{scope}** \u00b7 {sent_n} qids "
                f"({with_prior} expanded from prior AI Coach reports)._\n\n"
                f"_{appended_msg}_"
            )
            dlg = AICoachDialog(
                self, report_md + footer, num_questions=sent_n,
                model=f"{state.get('ai_backend_label', 'AI')} \u00b7 deep-review \u00b7 {scope_label}",
                master_path=deep_path,
            )
            dlg.setWindowTitle("Deep Review - per-question mini-class")
            self._show_coach_dialog_nonmodal(dlg)
            return

        if is_prebrief:
            scope = str(state.get("scope", ""))
            scope_label = str(state.get("scope_label", ""))
            sent_n = int(state.get("sent_n", 0) or 0)
            sent_qids = sorted(state.get("sent_qids", set()))

            # Pre-Brief writes into ai_coach_reports.md so Deep Review picks it
            # up later as `prior_report`. Marker scope tagged as `prebrief-<orig>`
            # so AI Coach dedup also treats these qids as already covered.
            provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
            master_path = self._provider_report_path("coach", provider_for_save)
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            qids_str = ", ".join(str(q) for q in sent_qids)
            tagged_scope = f"prebrief-{scope}"
            section = (
                f"\n\n<!-- AI_COACH_BATCH qids=[{qids_str}] scope={tagged_scope} ts={ts} -->\n"
                f"# {ts} - Pre-Brief - scope: {scope} - qids: {qids_str}\n\n"
                f"{report_md}\n"
            )
            try:
                with master_path.open("a", encoding="utf-8") as f:
                    f.write(section)
                appended_msg = f"Appended to: {master_path.name}"
            except Exception as _ex:
                appended_msg = f"(append failed: {_ex})"
            footer = (
                f"\n\n---\n_Pre-Brief saved \u00b7 scope **{scope}** \u00b7 {sent_n} qids. "
                f"Ahora corre Deep Review sobre los mismos qids para la c\u00e1tedra completa._\n\n"
                f"_{appended_msg}_"
            )
            dlg = AICoachDialog(
                self, report_md + footer, num_questions=sent_n,
                model=f"{state.get('ai_backend_label', 'AI')} \u00b7 pre-brief \u00b7 {scope_label}",
                master_path=master_path,
            )
            dlg.setWindowTitle("Pre-Brief - context dossier")
            self._show_coach_dialog_nonmodal(dlg)
            return

        scope = str(state.get("scope", ""))
        pool = list(state.get("pool", []))
        sent_n = int(state.get("sent_n", 0) or 0)
        scope_label = str(state.get("scope_label", ""))
        sent_qids = set(state.get("sent_qids", set()))

        self.ai_coach_sent.setdefault(scope, set()).update(sent_qids)
        self._save_ai_coach_sent()

        # Append to single master reports file (accumulates across sessions)
        provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
        master_path = self._provider_report_path("coach", provider_for_save)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        qids_str = ", ".join(str(q) for q in sorted(sent_qids))
        section = (
            f"\n\n<!-- AI_COACH_BATCH qids=[{qids_str}] scope={scope} ts={ts} -->\n"
            f"# {ts} - scope: {scope} - qids: {qids_str}\n\n"
            f"{report_md}\n"
        )
        try:
            with master_path.open("a", encoding="utf-8") as f:
                f.write(section)
            appended_msg = f"Appended to: {master_path.name}"
        except Exception as _ex:
            appended_msg = f"(append failed: {_ex})"

        remaining = len(pool) - len(self.ai_coach_sent[scope] & set(pool))
        footer = (
            f"\n\n---\n_Batch saved. Scope **{scope}**: "
            f"{len(self.ai_coach_sent[scope] & set(pool))}/{len(pool)} analyzed, "
            f"{remaining} pending. Click AI Coach again for the next batch._\n\n"
            f"_{appended_msg}_"
        )
        dlg = AICoachDialog(
            self, report_md + footer, num_questions=sent_n,
            model=f"{state.get('ai_backend_label', 'AI')} \u00b7 {scope_label}",
            master_path=master_path,
        )
        self._show_coach_dialog_nonmodal(dlg)


    def _load_flagged_bank(self) -> List[int]:
        raw = self.settings.value("bank/flagged", "[]", str) or "[]"
        try:
            data = json.loads(raw)
            if isinstance(data, list):
                return [int(x) for x in data if str(x).isdigit()]
        except Exception:
            pass
        return []

    def _save_flagged_bank(self) -> None:
        self.settings.setValue("bank/flagged", json.dumps(sorted(self.flagged)))
        self.settings.sync()


    def _load_q_stats(self) -> Dict[int, Dict[str, int]]:
        raw = self.settings.value("stats/per_question", "{}", str) or "{}"
        try:
            data = json.loads(raw)
            if isinstance(data, dict):
                out: Dict[int, Dict[str, int]] = {}
                for k, v in data.items():
                    try:
                        qid = int(k)
                    except Exception:
                        continue
                    if isinstance(v, dict):
                        out[qid] = {
                            "attempts": int(v.get("attempts", 0) or 0),
                            "wrong": int(v.get("wrong", 0) or 0),
                            "correct": int(v.get("correct", 0) or 0),
                        }
                return out
        except Exception:
            pass
        return {}

    def _save_q_stats(self) -> None:
        payload = {str(k): v for k, v in self.q_stats.items()}
        self.settings.setValue("stats/per_question", json.dumps(payload))
        self.settings.sync()


    def _load_agg_stats(self, key: str) -> Dict[str, Dict[str, int]]:
        raw = self.settings.value(key, "{}", str) or "{}"
        try:
            data = json.loads(raw)
            if isinstance(data, dict):
                out: Dict[str, Dict[str, int]] = {}
                for k, v in data.items():
                    if not k:
                        continue
                    if isinstance(v, dict):
                        out[str(k)] = {
                            "attempts": int(v.get("attempts", 0) or 0),
                            "wrong": int(v.get("wrong", 0) or 0),
                            "correct": int(v.get("correct", 0) or 0),
                        }
                return out
        except Exception:
            pass
        return {}

    def _save_agg_stats(self, key: str, data: Dict[str, Dict[str, int]]) -> None:
        payload = {str(k): v for k, v in data.items()}
        self.settings.setValue(key, json.dumps(payload, ensure_ascii=False))
        self.settings.sync()

    def _bump_agg(self, store: Dict[str, Dict[str, int]], key: str, correct: bool) -> None:
        k = (key or "").strip() or "(unlabeled)"
        s = store.get(k)
        if not s:
            s = {"attempts": 0, "wrong": 0, "correct": 0}
            store[k] = s
        s["attempts"] += 1
        if correct:
            s["correct"] += 1
        else:
            s["wrong"] += 1
    def _bump_stat(self, qid: int, correct: bool) -> None:
        # per-question stats
        s = self.q_stats.get(qid)
        if not s:
            s = {"attempts": 0, "wrong": 0, "correct": 0, "streak": 0}
            self.q_stats[qid] = s
        s["attempts"] += 1
        if correct:
            s["correct"] += 1
            s["streak"] = int(s.get("streak", 0) or 0) + 1
        else:
            s["wrong"] += 1
            s["streak"] = 0
        self._save_q_stats()

        # per-pattern / per-tag stats (best-effort)
        q = self._q_by_id(qid)
        if q:
            self._bump_agg(self.p_stats, q.pattern_id, correct)
            for t in (q.tags or []):
                self._bump_agg(self.t_stats, t, correct)
            self._save_agg_stats("stats/per_pattern", self.p_stats)
            self._save_agg_stats("stats/per_tag", self.t_stats)

        # SRS state (file: srs_state.json)
        try:
            srs = _load_json(_here("srs_state.json"), {}) or {}
            _srs_update(srs, qid, correct)
            _save_json_atomic(_here("srs_state.json"), srs)
        except Exception:
            pass

        # Concept mastery rollup (file: concept_mastery.json)
        try:
            if q:
                opts_text = " ".join((q.options or {}).values())
                concepts = _detect_concepts(f"{q.stem or ''} {opts_text}")
                if concepts:
                    cm = _load_json(_here("concept_mastery.json"), {}) or {}
                    _cm_bump(cm, concepts, qid, correct)
                    _save_json_atomic(_here("concept_mastery.json"), cm)
        except Exception:
            pass

        # Confidence log (file: confidence_log.jsonl) - append-only
        try:
            conf = "sure"
            cb = getattr(self, "cb_confidence", None)
            if cb is not None:
                conf = (cb.currentText() or "sure").lower().split()[0]
            ans = "".join(sorted(self.selections.get(qid, set()) or [])) if hasattr(self, "selections") else ""
            _conf_append({
                "ts": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
                "qid": int(qid), "confidence": conf, "correct": bool(correct), "ans": ans,
            })
        except Exception:
            pass

    def _snapshot_summary(self, snapshot: Dict[str, object]) -> str:
        mode = str(snapshot.get("mode") or "Unknown")
        order = snapshot.get("order_qids", [])
        total = len(order) if isinstance(order, list) else 0
        idx = int(snapshot.get("idx", 0) or 0)
        current_qid = snapshot.get("current_qid")
        q_text = f"Q{int(current_qid)}" if current_qid not in (None, "") else f"index {idx + 1}"
        answered = 0
        selections = snapshot.get("selections", {})
        if isinstance(selections, dict):
            answered = sum(1 for _, letters in selections.items() if isinstance(letters, (list, tuple, set)) and bool(letters))
        saved_at = str(snapshot.get("saved_at") or "").strip() or "(unknown time)"
        parts = [
            f"Mode: {mode}",
            f"Current: {q_text} ({idx + 1}/{total if total else '?'})",
            f"Answered: {answered}",
            f"Saved: {saved_at}",
        ]
        return "\n".join(parts)

    def _prompt_resume_autosave(self, snapshot: Dict[str, object]) -> bool:
        msg = QMessageBox(self)
        msg.setWindowTitle("Resume autosave?")
        msg.setIcon(QMessageBox.Question)
        msg.setText("An autosave session was found.")
        msg.setInformativeText(self._snapshot_summary(snapshot))
        resume_btn = msg.addButton("Resume", QMessageBox.ButtonRole.AcceptRole)
        discard_btn = msg.addButton("Discard", QMessageBox.ButtonRole.DestructiveRole)
        msg.setDefaultButton(resume_btn)
        msg.exec()
        return msg.clickedButton() is resume_btn

    def _prefer_newer_duplicate_docx(self, path: Path) -> Path:
        try:
            base = path.resolve()
        except Exception:
            base = path

        try:
            if base.parent.name.lower() != "00_active":
                return base
            promoted = base.parent.parent / base.name
            if promoted.exists() and promoted.stat().st_mtime > base.stat().st_mtime:
                return promoted.resolve()
        except Exception:
            pass
        return base

    def _restore_last(self) -> None:
        autosave = self._load_state_from_key(self._session_autosave_key)
        if autosave:
            snap_docx = str(autosave.get("docx_path") or "").strip()
            if snap_docx and Path(snap_docx).exists():
                if self._prompt_resume_autosave(autosave):
                    self._pending_restore_state = autosave
                    self._pending_restore_source = self._autosave_profile_name
                else:
                    self.settings.remove(self._session_autosave_key)
                    self.settings.remove(self._autosave_profile_key())
                    self.settings.sync()

        docx = self.settings.value("paths/docx", "", str) or ""
        if docx and Path(docx).exists():
            preferred_docx = self._prefer_newer_duplicate_docx(Path(docx))
            self.docx_path = preferred_docx
            if preferred_docx != Path(docx):
                self.settings.setValue("paths/docx", str(preferred_docx))
            self.lbl_doc.setText(f"DOCX: {self.docx_path}")

        ov = self.settings.value("paths/overrides", "", str) or ""
        if ov and Path(ov).exists():
            self._load_overrides_from_path(Path(ov))
            if self.questions:
                apply_patterns_from_overrides(self.questions, self.ov_patterns, self.ov_tags)

        self.mode = self.settings.value("quiz/mode", "Exam", str) or "Exam"
        self.time_limit_min = int(self.settings.value("quiz/minutes", 170) or 170)
        self.session_count = int(self.settings.value("quiz/count", 0) or 0)
        try:
            self.min_wrong = max(0, int(self.settings.value("quiz/min_wrong", self.min_wrong)))
        except Exception:
            self.min_wrong = max(0, int(self.min_wrong))
        self.custom_ids_text = self.settings.value("quiz/custom_ids", self.custom_ids_text, str) or self.custom_ids_text
        selected_profile = self.settings.value(self._session_profile_selected_key, "", str) or ""
        self.cb_mode.setCurrentText(self.mode)
        self.sp_time.setValue(self.time_limit_min)
        self.sp_count.setValue(self.session_count)
        self.sp_min_wrong.setValue(self.min_wrong)
        self.le_ids.setText(self.custom_ids_text)
        if selected_profile:
            self.cb_state_profile.setEditText(selected_profile)

        if self._pending_restore_state:
            snap_docx = str(self._pending_restore_state.get("docx_path") or "").strip()
            if snap_docx and Path(snap_docx).exists():
                self.docx_path = self._prefer_newer_duplicate_docx(Path(snap_docx))
                self.lbl_doc.setText(f"DOCX: {self.docx_path}")
                snap_ov = str(self._pending_restore_state.get("overrides_path") or "").strip()
                if snap_ov:
                    self.lbl_over.setText(f"Overrides: {snap_ov}")

        if self.docx_path and self.docx_path.exists():
            self._load_docx_async(self.docx_path)

    def _save_paths(self) -> None:
        if self.docx_path:
            self.settings.setValue("paths/docx", str(self.docx_path))
        else:
            self.settings.remove("paths/docx")
        if self.overrides_path:
            self.settings.setValue("paths/overrides", str(self.overrides_path))
        else:
            self.settings.remove("paths/overrides")
        self.settings.setValue("quiz/mode", self.mode)
        self.settings.setValue("quiz/minutes", int(self.time_limit_min))
        self.settings.setValue("quiz/count", int(self.session_count))
        self.settings.sync()

    def _profile_storage_key(self, name: str) -> str:
        raw = (name or "default").strip() or "default"
        enc = base64.urlsafe_b64encode(raw.encode("utf-8")).decode("ascii").rstrip("=")
        return f"session/profiles/{enc or 'default'}"

    def _is_reserved_profile_name(self, name: str) -> bool:
        return (name or "").strip().lower() in {"autosave", self._autosave_profile_name.lower()}

    def _autosave_profile_key(self) -> str:
        return self._profile_storage_key(self._autosave_profile_name)

    def _load_saved_snapshot(self) -> Optional[Dict[str, object]]:
        snapshot = self._load_state_from_key(self._autosave_profile_key())
        if snapshot:
            return snapshot
        return self._load_state_from_key(self._session_autosave_key)

    def _load_profile_names(self) -> List[str]:
        raw = self.settings.value(self._session_profile_names_key, "[]", str) or "[]"
        try:
            data = json.loads(raw)
            if isinstance(data, list):
                out: List[str] = []
                seen: Set[str] = set()
                for item in data:
                    name = str(item or "").strip()
                    if not name or name in seen or self._is_reserved_profile_name(name):
                        continue
                    seen.add(name)
                    out.append(name)
                return out
        except Exception:
            pass
        return []

    def _save_profile_names(self, names: List[str]) -> None:
        cleaned: List[str] = []
        seen: Set[str] = set()
        for item in names:
            name = str(item or "").strip()
            if not name or name in seen or self._is_reserved_profile_name(name):
                continue
            seen.add(name)
            cleaned.append(name)
        self.settings.setValue(self._session_profile_names_key, json.dumps(cleaned, ensure_ascii=False))
        self.settings.sync()

    def _refresh_profile_combo(self) -> None:
        if not hasattr(self, "cb_state_profile"):
            return
        current = self.cb_state_profile.currentText().strip() or self.settings.value(self._session_profile_selected_key, "", str) or ""
        names = self._load_profile_names()
        if self._load_saved_snapshot() and self._autosave_profile_name not in names:
            names = [self._autosave_profile_name] + names
        self.cb_state_profile.blockSignals(True)
        self.cb_state_profile.clear()
        self.cb_state_profile.addItems(names)
        if current:
            self.cb_state_profile.setEditText(current)
        elif names:
            preferred_idx = 0
            for idx, name in enumerate(names):
                if not self._is_reserved_profile_name(name):
                    preferred_idx = idx
                    break
            self.cb_state_profile.setCurrentIndex(preferred_idx)
        self.cb_state_profile.blockSignals(False)

    def _current_profile_name(self) -> str:
        if hasattr(self, "cb_state_profile"):
            try:
                txt = self.cb_state_profile.currentText().strip()
                if txt:
                    if txt.lower() == "autosave":
                        return self._autosave_profile_name
                    return txt
            except Exception:
                pass
        saved = self.settings.value(self._session_profile_selected_key, "", str) or ""
        saved_name = saved.strip()
        if saved_name.lower() == "autosave":
            return self._autosave_profile_name
        return saved_name or "default"

    def _session_is_restorable(self) -> bool:
        return bool(self.docx_path and self.questions and self.order)

    def _serialize_letter_map(self, data: Dict[int, Set[str]]) -> Dict[str, List[str]]:
        out: Dict[str, List[str]] = {}
        for qid, letters in (data or {}).items():
            try:
                qid_int = int(qid)
            except Exception:
                continue
            out[str(qid_int)] = sorted({str(ch).strip().upper() for ch in (letters or set()) if str(ch).strip()})
        return out

    def _serialize_frozen_map(self, data: Dict[int, FrozenSet[str]]) -> Dict[str, List[str]]:
        out: Dict[str, List[str]] = {}
        for qid, letters in (data or {}).items():
            try:
                qid_int = int(qid)
            except Exception:
                continue
            out[str(qid_int)] = sorted({str(ch).strip().upper() for ch in (letters or frozenset()) if str(ch).strip()})
        return out

    def _deserialize_letter_map(self, raw: object) -> Dict[int, Set[str]]:
        out: Dict[int, Set[str]] = {}
        if not isinstance(raw, dict):
            return out
        for k, v in raw.items():
            try:
                qid = int(k)
            except Exception:
                continue
            letters = set()
            if isinstance(v, (list, tuple, set)):
                for item in v:
                    ch = str(item or "").strip().upper()
                    if ch:
                        letters.add(ch)
            out[qid] = letters
        return out

    def _deserialize_frozen_map(self, raw: object) -> Dict[int, FrozenSet[str]]:
        out: Dict[int, FrozenSet[str]] = {}
        if not isinstance(raw, dict):
            return out
        for k, v in raw.items():
            try:
                qid = int(k)
            except Exception:
                continue
            letters: Set[str] = set()
            if isinstance(v, (list, tuple, set)):
                for item in v:
                    ch = str(item or "").strip().upper()
                    if ch:
                        letters.add(ch)
            out[qid] = frozenset(letters)
        return out

    def _serialize_option_order(self) -> Dict[str, List[str]]:
        out: Dict[str, List[str]] = {}
        for qid, letters in (self.option_order or {}).items():
            try:
                qid_int = int(qid)
            except Exception:
                continue
            out[str(qid_int)] = [str(ch).strip().upper() for ch in (letters or []) if str(ch).strip()]
        return out

    def _current_session_state(self, profile_name: str = "") -> Dict[str, object]:
        current_qid: Optional[int] = None
        if self.questions and self.order and 0 <= self.idx < len(self.order):
            try:
                current_qid = int(self.questions[self.order[self.idx]].qid)
            except Exception:
                current_qid = None
        return {
            "version": 1,
            "profile_name": (profile_name or "").strip(),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "docx_path": str(self.docx_path) if self.docx_path else "",
            "docx_signature": self.docx_signature or (str(self.docx_path) and _compute_file_sha256(self.docx_path) if self.docx_path else ""),
            "bank_signature": self.bank_signature or _compute_questions_signature(self.questions),
            "overrides_path": str(self.overrides_path) if self.overrides_path else "",
            "mode": self.mode,
            "time_limit_min": int(self.time_limit_min),
            "session_count": int(self.session_count),
            "min_wrong": int(self.min_wrong),
            "custom_ids_text": self.custom_ids_text,
            "shuffle_questions": bool(self.shuffle_questions),
            "shuffle_options": bool(self.shuffle_options),
            "auto_check_next_enabled": bool(self.auto_check_next_enabled),
            "bank_text": self.cb_bank.currentText() if hasattr(self, "cb_bank") else "All questions",
            "order_qids": [int(self.questions[i].qid) for i in (self.order or [])],
            "idx": int(self.idx),
            "current_qid": current_qid,
            "selections": self._serialize_letter_map(self.selections),
            "option_order": self._serialize_option_order(),
            "submitted": bool(self.submitted),
            "seconds_left": None if self.seconds_left is None else int(self.seconds_left),
            "seconds_spent": int(self.seconds_spent),
            "practice_correct": sorted(int(x) for x in self.practice_correct),
            "practice_attempted": sorted(int(x) for x in self.practice_attempted),
            "last_incorrect": sorted(int(x) for x in self.last_incorrect),
            "session_new_initial": sorted(int(x) for x in self.session_new_initial),
            "session_new_remaining": sorted(int(x) for x in self.session_new_remaining),
            "session_results_recorded": bool(self._session_results_recorded),
            "last_checked_sel": self._serialize_frozen_map(self._last_checked_sel),
            "last_scored_sel": self._serialize_frozen_map(self._last_scored_sel),
            "feedback_visible": bool(self.txt_feedback.isVisible()) if hasattr(self, "txt_feedback") else False,
            "feedback_html": self.txt_feedback.toHtml() if hasattr(self, "txt_feedback") and self.txt_feedback.isVisible() else "",
        }

    def _save_state_to_key(self, key: str, profile_name: str = "", silent: bool = True) -> bool:
        if not self._session_is_restorable():
            if not silent:
                QMessageBox.information(self, "No active session", "Start a session first so there is progress to save.")
            return False
        snapshot = self._current_session_state(profile_name=profile_name)
        self.settings.setValue(key, json.dumps(snapshot, ensure_ascii=False))
        self.settings.sync()
        return True

    def _load_state_from_key(self, key: str) -> Optional[Dict[str, object]]:
        raw = self.settings.value(key, "", str) or ""
        if not raw:
            return None
        try:
            data = json.loads(raw)
        except Exception:
            return None
        return data if isinstance(data, dict) else None

    def _autosave_status_text(self, when_text: str) -> str:
        return f"Autosave: {when_text}"

    def _set_autosave_label(self, when_text: str) -> None:
        if hasattr(self, "lbl_autosave"):
            self.lbl_autosave.setText(self._autosave_status_text(when_text))

    def _schedule_autosave(self) -> None:
        if not self._session_is_restorable():
            self._autosave_debounce.stop()
            self._autosave_periodic.stop()
            self._set_autosave_label("waiting for a session")
            return
        self._autosave_debounce.start()
        if not self._autosave_periodic.isActive():
            self._autosave_periodic.start()

    def _save_autosave(self) -> None:
        if not self._session_is_restorable():
            self._autosave_periodic.stop()
            self._set_autosave_label("waiting for a session")
            return
        saved = self._save_state_to_key(self._session_autosave_key, profile_name=self._autosave_profile_name, silent=True)
        saved = self._save_state_to_key(self._autosave_profile_key(), profile_name=self._autosave_profile_name, silent=True) or saved
        if saved:
            self._refresh_profile_combo()
            self._set_autosave_label(f"saved {datetime.now().strftime('%H:%M:%S')}")

    def save_named_state(self) -> None:
        name = self._current_profile_name()
        if not name:
            name = "default"
        if self._is_reserved_profile_name(name):
            QMessageBox.information(
                self,
                "Reserved profile",
                f"The profile name '{self._autosave_profile_name}' is reserved for autosave. Choose another name for manual saves.",
            )
            return
        if not self._save_state_to_key(self._profile_storage_key(name), profile_name=name, silent=False):
            return
        names = self._load_profile_names()
        if name not in names:
            names.append(name)
            self._save_profile_names(names)
        self.settings.setValue(self._session_profile_selected_key, name)
        self.settings.sync()
        self._refresh_profile_combo()
        self._set_autosave_label(f"manual save '{name}' at {datetime.now().strftime('%H:%M:%S')}")
        QMessageBox.information(self, "State saved", f"Saved session profile '{name}'.")

    def _paths_match_snapshot(self, snapshot: Dict[str, object]) -> bool:
        snap_docx = str(snapshot.get("docx_path") or "").strip()
        snap_overrides = str(snapshot.get("overrides_path") or "").strip()
        cur_docx = str(self.docx_path) if self.docx_path else ""
        cur_overrides = str(self.overrides_path) if self.overrides_path else ""
        try:
            docx_ok = bool(cur_docx and snap_docx) and Path(cur_docx).resolve() == Path(snap_docx).resolve()
        except Exception:
            docx_ok = (cur_docx == snap_docx)
        try:
            overrides_ok = (not cur_overrides and not snap_overrides) or (
                bool(cur_overrides and snap_overrides) and Path(cur_overrides).resolve() == Path(snap_overrides).resolve()
            )
        except Exception:
            overrides_ok = (cur_overrides == snap_overrides)
        return docx_ok and overrides_ok and self._snapshot_matches_loaded_bank(snapshot)

    def _snapshot_order_qids(self, snapshot: Dict[str, object]) -> Set[int]:
        raw_order = snapshot.get("order_qids", [])
        qids: Set[int] = set()
        if isinstance(raw_order, list):
            for item in raw_order:
                try:
                    qids.add(int(item))
                except Exception:
                    continue
        return qids

    def _loaded_qids(self) -> Set[int]:
        qids: Set[int] = set()
        for q in self.questions or []:
            try:
                qids.add(int(q.qid))
            except Exception:
                continue
        return qids

    def _docx_path_matches_snapshot(self, snapshot: Dict[str, object]) -> bool:
        snap_docx = str(snapshot.get("docx_path") or "").strip()
        cur_docx = str(self.docx_path) if self.docx_path else ""
        if not snap_docx or not cur_docx:
            return False
        try:
            return Path(cur_docx).resolve() == Path(snap_docx).resolve()
        except Exception:
            return cur_docx == snap_docx

    def _docx_name_matches_snapshot(self, snapshot: Dict[str, object]) -> bool:
        snap_docx = str(snapshot.get("docx_path") or "").strip()
        cur_docx = str(self.docx_path) if self.docx_path else ""
        if not snap_docx or not cur_docx:
            return False
        try:
            return Path(cur_docx).name.lower() == Path(snap_docx).name.lower()
        except Exception:
            return False

    def _snapshot_qids_exist_in_loaded_bank(self, snapshot: Dict[str, object]) -> bool:
        saved_qids = self._snapshot_order_qids(snapshot)
        loaded_qids = self._loaded_qids()
        return bool(saved_qids and loaded_qids and saved_qids.issubset(loaded_qids))

    def _snapshot_bank_compatible_by_qids(self, snapshot: Dict[str, object]) -> bool:
        if not self._snapshot_qids_exist_in_loaded_bank(snapshot):
            return False
        if self._docx_path_matches_snapshot(snapshot) or self._docx_name_matches_snapshot(snapshot):
            return True
        snap_docx_sig = str(snapshot.get("docx_signature") or "").strip()
        if snap_docx_sig and self.docx_signature:
            return snap_docx_sig == self.docx_signature
        return False

    def _snapshot_matches_loaded_bank(self, snapshot: Dict[str, object]) -> bool:
        snap_docx_sig = str(snapshot.get("docx_signature") or "").strip()
        snap_bank_sig = str(snapshot.get("bank_signature") or "").strip()
        if snap_bank_sig and self.bank_signature and snap_bank_sig != self.bank_signature:
            return self._snapshot_bank_compatible_by_qids(snapshot)
        if snap_docx_sig and self.docx_signature and snap_docx_sig != self.docx_signature:
            return self._snapshot_bank_compatible_by_qids(snapshot)
        return True

    def _request_restore_state(self, snapshot: Optional[Dict[str, object]], source: str, silent: bool = False) -> bool:
        if not snapshot:
            if not silent:
                QMessageBox.information(self, "No saved state", f"There is no saved state for {source}.")
            return False

        snap_docx = str(snapshot.get("docx_path") or "").strip()
        if not snap_docx:
            if not silent:
                QMessageBox.warning(self, "Invalid state", "The saved state does not include a DOCX path.")
            return False

        path = self._prefer_newer_duplicate_docx(Path(snap_docx))
        if not path.exists():
            if not silent:
                QMessageBox.warning(self, "Missing DOCX", f"The saved DOCX was not found:\n{path}")
            return False

        self._pending_restore_state = snapshot
        self._pending_restore_source = source

        if self.questions and self._paths_match_snapshot(snapshot):
            ok = self._restore_pending_state()
            if ok and not silent:
                QMessageBox.information(self, "State restored", f"Restored {source}.")
            return ok

        self.docx_path = path
        self.settings.setValue("paths/docx", str(path))
        self.lbl_doc.setText(f"DOCX: {self.docx_path}")
        self._load_docx_async(path)
        return True

    def load_named_state(self) -> None:
        name = self._current_profile_name()
        if not name:
            QMessageBox.information(self, "Profile name", "Type or select a profile name first.")
            return
        if self._is_reserved_profile_name(name):
            snapshot = self._load_saved_snapshot()
        else:
            snapshot = self._load_state_from_key(self._profile_storage_key(name))
        if not snapshot:
            QMessageBox.information(self, "Profile not found", f"No saved state exists for '{name}'.")
            return
        if not self._is_reserved_profile_name(name):
            self.settings.setValue(self._session_profile_selected_key, name)
            self.settings.sync()
        self._refresh_profile_combo()
        self._request_restore_state(snapshot, f"profile '{name}'", silent=False)

    def delete_named_state(self) -> None:
        name = self._current_profile_name()
        if not name:
            QMessageBox.information(self, "Profile name", "Type or select a profile name first.")
            return
        if self._is_reserved_profile_name(name):
            self.settings.remove(self._session_autosave_key)
            self.settings.remove(self._autosave_profile_key())
            self.settings.sync()
        else:
            key = self._profile_storage_key(name)
            if self.settings.value(key, "", str):
                self.settings.remove(key)
            names = [item for item in self._load_profile_names() if item != name]
            self._save_profile_names(names)
        self._refresh_profile_combo()
        self._set_autosave_label(f"deleted profile '{name}'")
        QMessageBox.information(self, "Deleted", f"Deleted saved state '{name}'.")

    def _overrides_file_stamp(self, path: Optional[Path]) -> Optional[Tuple[int, int]]:
        if not path:
            return None
        try:
            st = path.stat()
            return int(st.st_mtime_ns), int(st.st_size)
        except Exception:
            return None

    def _load_overrides_from_path(self, path: Path) -> None:
        self.overrides_path = Path(path)
        self.overrides, self.notes, self.my_expl, self.ov_patterns, self.ov_tags = load_overrides_csv(self.overrides_path)
        self.overrides_stamp = self._overrides_file_stamp(self.overrides_path)
        self.lbl_over.setText(f"Overrides: {self.overrides_path} ({len(self.overrides)} fixes)")
        if self.questions and (self.ov_patterns or self.ov_tags):
            apply_patterns_from_overrides(self.questions, self.ov_patterns, self.ov_tags)

    def _clear_overrides(self) -> None:
        self.overrides_path = None
        self.overrides_stamp = None
        self.overrides = {}
        self.notes = {}
        self.my_expl = {}
        self.ov_patterns = {}
        self.ov_tags = {}
        self.lbl_over.setText("Overrides: (none)")

    def _maybe_reload_overrides(self) -> bool:
        """Reload the overrides CSV if it was edited outside the app."""
        path = self.overrides_path
        if not path:
            return False
        try:
            if not path.exists():
                return False
        except Exception:
            return False
        stamp = self._overrides_file_stamp(path)
        if stamp is None or stamp == self.overrides_stamp:
            return False
        self._load_overrides_from_path(path)
        try:
            self.lbl_hint.setText(f"Overrides reloaded: {len(self.overrides)} fixes")
        except Exception:
            pass
        return True

    def _apply_overrides_snapshot(self, snapshot: Dict[str, object]) -> None:
        ov_text = str(snapshot.get("overrides_path") or "").strip()
        if ov_text and Path(ov_text).exists():
            self._load_overrides_from_path(Path(ov_text))
        else:
            self._clear_overrides()

        if self.ov_patterns or self.ov_tags:
            apply_patterns_from_overrides(self.questions, self.ov_patterns, self.ov_tags)

    def _restore_pending_state(self) -> bool:
        snapshot = self._pending_restore_state
        source = self._pending_restore_source or "saved state"
        self._pending_restore_state = None
        self._pending_restore_source = ""
        if not snapshot:
            return False
        try:
            if not self._snapshot_matches_loaded_bank(snapshot):
                if source in {"autosave", self._autosave_profile_name}:
                    self.settings.remove(self._session_autosave_key)
                    self.settings.remove(self._autosave_profile_key())
                    self.settings.sync()
                QMessageBox.warning(
                    self,
                    "Saved state mismatch",
                    (
                        f"Could not restore {source} because the loaded question bank does not match the one "
                        "used when that session was saved.\n\nLoad the original DOCX or start a fresh session."
                    ),
                )
                return False
            self._apply_overrides_snapshot(snapshot)
            ok = self._apply_session_state(snapshot)
            if ok:
                self._set_autosave_label(f"restored {source}")
            return ok
        except Exception:
            traceback.print_exc()
            return False

    def _apply_session_state(self, snapshot: Dict[str, object]) -> bool:
        if not self.questions:
            return False

        mode = str(snapshot.get("mode") or "Exam")
        bank_text = str(snapshot.get("bank_text") or "All questions")
        time_limit = int(snapshot.get("time_limit_min", self.time_limit_min) or self.time_limit_min)
        session_count = int(snapshot.get("session_count", self.session_count) or 0)
        try:
            min_wrong = int(snapshot.get("min_wrong", self.min_wrong))
        except Exception:
            min_wrong = int(self.min_wrong)
        custom_ids_text = str(snapshot.get("custom_ids_text") or "")
        shuffle_questions = bool(snapshot.get("shuffle_questions", True))
        shuffle_options = bool(snapshot.get("shuffle_options", True))
        auto_check_next = bool(snapshot.get("auto_check_next_enabled", False))

        self.cb_mode.blockSignals(True)
        self.cb_bank.blockSignals(True)
        self.sp_time.blockSignals(True)
        self.sp_count.blockSignals(True)
        self.sp_min_wrong.blockSignals(True)
        self.chk_shuffle_q.blockSignals(True)
        self.chk_shuffle_o.blockSignals(True)
        self.chk_auto_next.blockSignals(True)
        self.cb_mode.setCurrentText(mode if mode in ("Exam", "Practice") else "Exam")
        idx_bank = self.cb_bank.findText(bank_text)
        if idx_bank >= 0:
            self.cb_bank.setCurrentIndex(idx_bank)
        self.sp_time.setValue(max(10, time_limit))
        self.sp_count.setValue(max(0, session_count))
        self.sp_min_wrong.setValue(max(0, min_wrong))
        self.chk_shuffle_q.setChecked(bool(shuffle_questions))
        self.chk_shuffle_o.setChecked(bool(shuffle_options))
        self.chk_auto_next.setChecked(bool(auto_check_next))
        self.cb_mode.blockSignals(False)
        self.cb_bank.blockSignals(False)
        self.sp_time.blockSignals(False)
        self.sp_count.blockSignals(False)
        self.sp_min_wrong.blockSignals(False)
        self.chk_shuffle_q.blockSignals(False)
        self.chk_shuffle_o.blockSignals(False)
        self.chk_auto_next.blockSignals(False)

        self.custom_ids_text = custom_ids_text
        self.le_ids.setText(custom_ids_text)
        restored_profile_name = str(snapshot.get("profile_name") or "").strip()
        if restored_profile_name and not self._is_reserved_profile_name(restored_profile_name):
            self.settings.setValue(self._session_profile_selected_key, restored_profile_name)
        else:
            current_name = self._current_profile_name()
            if current_name and not self._is_reserved_profile_name(current_name):
                self.settings.setValue(self._session_profile_selected_key, current_name)

        self.mode = self.cb_mode.currentText() or mode
        self.time_limit_min = max(10, time_limit)
        self.session_count = max(0, session_count)
        self.min_wrong = max(0, min_wrong)
        self.auto_check_next_enabled = bool(auto_check_next)
        self._sync_shuffle_flags()
        self.custom_ids_active = bool(self._parse_custom_ids())
        self._set_mode(self.mode)

        qid_to_index = {int(q.qid): idx for idx, q in enumerate(self.questions)}
        raw_order = snapshot.get("order_qids", [])
        order_qids: List[int] = []
        if isinstance(raw_order, list):
            for item in raw_order:
                try:
                    qid = int(item)
                except Exception:
                    continue
                if qid in qid_to_index:
                    order_qids.append(qid)
        if not order_qids:
            return False

        self.order = [qid_to_index[qid] for qid in order_qids]
        self.submitted = bool(snapshot.get("submitted", False))
        self.selections = {
            qid: letters
            for qid, letters in self._deserialize_letter_map(snapshot.get("selections", {})).items()
            if qid in qid_to_index
        }
        self.practice_correct = {qid for qid in snapshot.get("practice_correct", []) if qid in qid_to_index} if isinstance(snapshot.get("practice_correct", []), list) else set()
        self.practice_attempted = {qid for qid in snapshot.get("practice_attempted", []) if qid in qid_to_index} if isinstance(snapshot.get("practice_attempted", []), list) else set()
        self.last_incorrect = {qid for qid in snapshot.get("last_incorrect", []) if qid in qid_to_index} if isinstance(snapshot.get("last_incorrect", []), list) else set()
        self.session_new_initial = {qid for qid in snapshot.get("session_new_initial", []) if qid in qid_to_index} if isinstance(snapshot.get("session_new_initial", []), list) else set()
        self.session_new_remaining = {qid for qid in snapshot.get("session_new_remaining", []) if qid in qid_to_index} if isinstance(snapshot.get("session_new_remaining", []), list) else set()
        self._session_results_recorded = bool(snapshot.get("session_results_recorded", False))
        self._last_checked_sel = {
            qid: letters
            for qid, letters in self._deserialize_frozen_map(snapshot.get("last_checked_sel", {})).items()
            if qid in qid_to_index
        }
        self._last_scored_sel = {
            qid: letters
            for qid, letters in self._deserialize_frozen_map(snapshot.get("last_scored_sel", {})).items()
            if qid in qid_to_index
        }

        raw_opt = snapshot.get("option_order", {})
        self.option_order = {}
        for q in self.questions:
            default_letters = list(q.options.keys())
            saved_letters: List[str] = []
            if isinstance(raw_opt, dict):
                val = raw_opt.get(str(q.qid), [])
                if isinstance(val, list):
                    saved_letters = [str(ch).strip().upper() for ch in val if str(ch).strip()]
            self.option_order[q.qid] = saved_letters if saved_letters and set(saved_letters) == set(default_letters) else default_letters

        current_qid = snapshot.get("current_qid")
        idx = int(snapshot.get("idx", 0) or 0)
        if current_qid is not None:
            try:
                qid_target = int(current_qid)
                idx = order_qids.index(qid_target)
            except Exception:
                pass
        self.idx = min(max(0, idx), max(0, len(self.order) - 1))

        self._tick.stop()
        self.seconds_spent = max(0, int(snapshot.get("seconds_spent", 0) or 0))
        saved_seconds_left = snapshot.get("seconds_left", None)
        if self.mode == "Exam":
            if saved_seconds_left is None or str(saved_seconds_left) == "":
                self.seconds_left = max(0, self.time_limit_min * 60 - self.seconds_spent)
            else:
                self.seconds_left = max(0, int(saved_seconds_left))
            self.lbl_timer.setText(f"Time: {_fmt_mmss(self.seconds_left)}")
            if not self.submitted and not self._session_results_recorded and self.seconds_left > 0:
                self._tick.start(1000)
        else:
            self.seconds_left = None
            self.lbl_timer.setText("Time: --:--")
            if not self._session_results_recorded:
                self._tick.start(1000)

        self._populate_nav()
        self._render_current(keep_feedback=bool(snapshot.get("feedback_visible", False)))
        feedback_html = str(snapshot.get("feedback_html") or "")
        if bool(snapshot.get("feedback_visible", False)) and feedback_html and self.mode == "Practice":
            self.txt_feedback.setVisible(True)
            self.txt_feedback.setHtml(feedback_html)

        self._save_paths()
        self._schedule_autosave()
        return True

    # ---------- mode & timer ----------
    def _set_mode(self, mode: str) -> None:
        self.mode = mode
        is_exam = self.mode == "Exam"

        # Stop any running countdown when switching modes to avoid 'practice timer' confusion
        self._tick.stop()
        self.seconds_left = None

        self.sp_time.setEnabled(is_exam)
        self.btn_check.setVisible(not is_exam)
        try:
            self.chk_auto_next.setVisible(not is_exam)
        except Exception:
            pass

        # Use the same button for both modes
        self.btn_submit.setEnabled(True)
        self.btn_submit.setText("Submit (Exam)" if is_exam else "Finish (Practice)")

        self.lbl_timer.setText(f"\u23f1 {_fmt_mmss(self.time_limit_min*60)}" if is_exam else "\u23f1 --:--")
        self.txt_feedback.setVisible(False)
        self._save_paths()
        self._schedule_autosave()

    def _set_time_limit(self, minutes: int) -> None:
        self.time_limit_min = int(minutes)
        if self.mode == "Exam":
            self.lbl_timer.setText(f"\u23f1 {_fmt_mmss(self.time_limit_min*60)}")
        self._save_paths()
        self._schedule_autosave()

    def _set_session_count(self, count: int) -> None:
        # 0 = ALL
        try:
            self.session_count = max(0, int(count))
        except Exception:
            self.session_count = 0
        self.settings.setValue("quiz/count", int(self.session_count))
        self.settings.sync()
        self._schedule_autosave()

    def _set_min_wrong(self, v: int) -> None:
        try:
            self.min_wrong = max(0, int(v))
        except Exception:
            self.min_wrong = 0
        # AUTO-SWITCH bank to Wrong >= Min so the threshold actually applies.
        try:
            if hasattr(self, 'cb_bank'):
                cur = (self.cb_bank.currentText() or '')
                if not cur.startswith('Wrong'):
                    idx = self.cb_bank.findText('Wrong >= Min')
                    if idx < 0:
                        idx = self.cb_bank.findText('Wrong \u2265 Min')
                    if idx >= 0:
                        self.cb_bank.blockSignals(True)
                        self.cb_bank.setCurrentIndex(idx)
                        self.cb_bank.blockSignals(False)
        except Exception:
            pass
        self.settings.setValue("quiz/min_wrong", int(self.min_wrong))
        self.settings.sync()
        self._schedule_autosave()

    def _set_auto_check_next(self, state: int) -> None:
        self.auto_check_next_enabled = bool(int(state) != 0)
        self.settings.setValue("quiz/auto_check_next", 1 if self.auto_check_next_enabled else 0)
        self.settings.sync()
        self._schedule_autosave()


    def _set_custom_ids_text(self) -> None:
        try:
            txt = (self.le_ids.text() if hasattr(self, "le_ids") else self.custom_ids_text) or ""
            self.custom_ids_text = txt.strip()
        except Exception:
            self.custom_ids_text = ""
        self.settings.setValue("quiz/custom_ids", self.custom_ids_text)
        self.settings.sync()
        self._schedule_autosave()

    def _set_custom_ids_value(self, text: str) -> None:
        clean = (text or "").strip()
        if hasattr(self, "le_ids"):
            self.le_ids.blockSignals(True)
            self.le_ids.setText(clean)
            self.le_ids.blockSignals(False)
        self.custom_ids_text = clean
        self.settings.setValue("quiz/custom_ids", self.custom_ids_text)
        self.settings.sync()

    def _set_session_count_value(self, count: int) -> None:
        try:
            clean = max(0, int(count))
        except Exception:
            clean = 0
        if hasattr(self, "sp_count"):
            self.sp_count.blockSignals(True)
            self.sp_count.setValue(clean)
            self.sp_count.blockSignals(False)
        self.session_count = clean
        self.settings.setValue("quiz/count", int(self.session_count))
        self.settings.sync()

    def _start_bank_preset(self, bank_text: str, *, count: Optional[int] = None) -> None:
        if not self.questions:
            self.lbl_hint.setText("Load your DOCX to begin.")
            return
        self._set_custom_ids_value("")
        if count is not None:
            self._set_session_count_value(count)
        idx = self.cb_bank.findText(bank_text)
        if idx >= 0:
            self.cb_bank.blockSignals(True)
            self.cb_bank.setCurrentIndex(idx)
            self.cb_bank.blockSignals(False)
        self.start_session()

    def start_new_mix_session(self) -> None:
        self._start_bank_preset("New Qs + random mix")

    def start_new_only_session(self) -> None:
        self._start_bank_preset("New Qs only (306-322)", count=0)

    def _parse_custom_ids(self) -> List[int]:
        raw = ""
        try:
            raw = (self.le_ids.text() if hasattr(self, "le_ids") else self.custom_ids_text) or ""
        except Exception:
            raw = self.custom_ids_text or ""
        raw = raw.strip()
        if not raw:
            return []
        parts = re.split(r"[,\s]+", raw)
        out: List[int] = []
        seen: Set[int] = set()
        for p in parts:
            p = (p or "").strip()
            if not p:
                continue
            if not p.isdigit():
                continue
            qid = int(p)
            if qid in seen:
                continue
            seen.add(qid)
            out.append(qid)
        return out

    def _question_indices_by_custom_id_list(self, qids: List[int]) -> List[int]:
        m = {q.qid: i for i, q in enumerate(self.questions)}
        out: List[int] = []
        for qid in qids:
            if qid in m:
                out.append(m[qid])
        return out

    def _on_tick(self) -> None:
        self.seconds_spent += 1
        if self.seconds_left is not None:
            self.seconds_left -= 1
            if self.seconds_left <= 0:
                self.seconds_left = 0
                self._tick.stop()
                self.lbl_timer.setText("\u23f1 00:00")
                if not self.submitted and self.questions:
                    QMessageBox.information(self, "Time", "Time is up. Auto-submitting.")
                    self.submit_exam()
                return
            self.lbl_timer.setText(f"\u23f1 {_fmt_mmss(self.seconds_left)}")

    # ---------- loading ----------
    def pick_docx(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select DOCX", filter="Word Document (*.docx)")
        if path:
            self.docx_path = Path(path)
            self.lbl_doc.setText(f"DOCX: {self.docx_path}")
            self._save_paths()
            self._load_docx_async(self.docx_path)

    def _load_docx_async(self, path: Path) -> None:
        if self._thread and self._thread.isRunning():
            QMessageBox.warning(self, "Busy", "A DOCX is already loading.")
            return
        self._autosave_debounce.stop()
        self._autosave_periodic.stop()
        self.lbl_hint.setText("Loading DOCX...")
        self.list_q.clear()
        self.questions = []
        self.docx_signature = ""
        self.bank_signature = ""
        self.order = []
        self.idx = 0
        self.selections = {}
        # keep flagged bank across sessions
        self.submitted = False
        self.last_incorrect = set()
        self.practice_correct = set()
        self.practice_attempted = set()
        self._session_scored = set()
        self._last_scored_sel = {}
        self.seconds_left = None
        self.seconds_spent = 0
        self.lbl_progress.setText("0 / 0")
        self.txt_feedback.setVisible(False)

        self._thread = QThread(self)
        self._worker = LoadWorker(path)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(lambda p: self.lbl_hint.setText(f"Loading DOCX... {p}%"))
        self._worker.finished.connect(self._docx_loaded)
        self._thread.start()

    @Slot(object, object)
    def _docx_loaded(self, questions: object, error: object) -> None:
        if self._thread:
            self._thread.quit()
            self._thread.wait()
        self._thread = None
        if self._worker:
            self._worker.deleteLater()
        self._worker = None

        if error:
            QMessageBox.critical(self, "DOCX Error", str(error))
            self.lbl_hint.setText("Failed to load DOCX.")
            return

        self.questions = list(questions) if questions else []
        self.docx_signature = _compute_file_sha256(self.docx_path) if self.docx_path else ""
        self.bank_signature = _compute_questions_signature(self.questions)
        if hasattr(self, 'sp_count'):
            try:
                self.sp_count.setMaximum(max(0, len(self.questions)))
            except Exception:
                pass
        if not self.questions:
            QMessageBox.warning(self, "Empty", "No questions were parsed from this DOCX.")
            self.lbl_hint.setText("No questions found.")
            return

        # Auto-label patterns (best-effort). You can override via Overrides CSV.
        for q in self.questions:
            if not q.pattern_id and not q.tags:
                pid, tags = infer_pattern_and_tags(q)
                q.pattern_id = pid
                q.tags = tags
        if self.ov_patterns or self.ov_tags:
            apply_patterns_from_overrides(self.questions, self.ov_patterns, self.ov_tags)

        # Prune flagged IDs that don't exist in the current bank
        try:
            valid_ids = {qq.qid for qq in self.questions}
            self.flagged = {qid for qid in self.flagged if qid in valid_ids}
            self._save_flagged_bank()
        except Exception:
            pass

        self.lbl_hint.setText(f"Loaded {len(self.questions)} questions.")
        if self._pending_restore_state and self._restore_pending_state():
            return
        self.start_session()

    def pick_overrides(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select Overrides CSV", filter="CSV Files (*.csv)")
        if path:
            self._load_overrides_from_path(Path(path))
            self._save_paths()
            self._refresh_nav()
            self._render_current()

    def export_template(self) -> None:
        if not self.questions:
            QMessageBox.warning(self, "No bank", "Load a DOCX first.")
            return
        base = "ans_c01_overrides_template.csv"
        if self.docx_path:
            base = self.docx_path.with_suffix("").name + "_overrides_template.csv"
        path, _ = QFileDialog.getSaveFileName(self, "Save Overrides Template", base, "CSV Files (*.csv)")
        if path:
            export_overrides_template(Path(path), self.questions)
            QMessageBox.information(self, "Saved", f"Template saved:\n{Path(path).resolve()}")

    # ---------- session construction ----------
    def _question_indices_by_qids(self, qids: Set[int]) -> List[int]:
        m = {q.qid: i for i, q in enumerate(self.questions)}
        out = [m[qid] for qid in qids if qid in m]
        out.sort(key=lambda i: self.questions[i].qid)
        return out

    def _new_question_qids(self) -> Set[int]:
        return {int(q.qid) for q in self.questions if NEW_QID_START <= int(q.qid) <= NEW_QID_END}

    def _weighted_wrong_order(self, qids: Set[int]) -> List[int]:
        """Order wrong-mode qids by struggle, not uniform random.

        Score = (wrong + lapses + 1) * streak_decay + hours_since_review/24 + jitter
          - wrong: q_stats[qid]['wrong']
          - lapses, last_reviewed: srs_state.json
          - streak_decay = 1 - 0.25*min(streak,3) so questions you've gotten right
            twice in a row weigh less than fresh misses
          - hours_since: never-reviewed treated as 720h (30d) so new misses surface
          - jitter: 0..0.4 to avoid identical-score determinism

        Leeches (lapses>=3 AND streak<2) get pinned to the front, capped at 5.
        """
        if not qids:
            return []
        srs = _load_json(_here("srs_state.json"), {}) or {}
        now = datetime.now()
        scored: List[tuple] = []
        leeches: List[tuple] = []
        for qid in qids:
            s = self.q_stats.get(qid, {}) or {}
            wrong_n = int(s.get("wrong", 0) or 0)
            streak = int(s.get("streak", 0) or 0)
            e = srs.get(str(qid)) or {}
            lapses = int(e.get("lapses", 0) or 0)
            last = str(e.get("last_reviewed", "") or "")
            if last:
                try:
                    d = datetime.fromisoformat(last)
                    hours_since = max(0.0, (now - d).total_seconds() / 3600.0)
                except Exception:
                    hours_since = 720.0
            else:
                hours_since = 720.0
            decay = 1.0 - 0.25 * min(streak, 3)
            score = (wrong_n + lapses + 1) * decay + hours_since / 24.0
            score += random.uniform(0.0, 0.4)
            entry = (score, qid)
            if lapses >= 3 and streak < 2:
                leeches.append(entry)
            else:
                scored.append(entry)
        leeches.sort(key=lambda t: -t[0])
        scored.sort(key=lambda t: -t[0])
        ordered_qids = [q for _, q in leeches[:5]] + [q for _, q in leeches[5:]] + [q for _, q in scored]
        m = {q.qid: i for i, q in enumerate(self.questions)}
        return [m[q] for q in ordered_qids if q in m]

    def _build_order(self) -> List[int]:
        if not self.questions:
            return []

        # SRS "Due Today" one-shot override (set by start_due_today, consumed once)
        due_override = getattr(self, "_due_today_override", None)
        if due_override:
            self._due_today_override = None
            qids = set(due_override)
            idxs = self._weighted_wrong_order(qids) if qids else []
            if not idxs:
                return list(range(len(self.questions)))
            return idxs

        # Custom IDs override bank selection (comma-separated list in the IDs box)
        custom_list = self._parse_custom_ids()
        if custom_list:
            idxs = self._question_indices_by_custom_id_list(custom_list)
            if not idxs:
                return list(range(len(self.questions)))
            if self.shuffle_questions:
                self._rng.shuffle(idxs)
            return idxs

        pick = self.cb_bank.currentText()
        weighted_mode = False  # set True by wrong-related branches; suppresses trailing shuffle

        if pick.startswith("New Qs + random mix"):
            new_qids = self._new_question_qids()
            if not new_qids:
                idxs = list(range(len(self.questions)))
                if self.shuffle_questions:
                    self._rng.shuffle(idxs)
                return idxs

            new_idxs = self._question_indices_by_qids(new_qids)
            other_idxs = [idx for idx, q in enumerate(self.questions) if int(q.qid) not in new_qids]
            if self.shuffle_questions:
                self._rng.shuffle(new_idxs)
                self._rng.shuffle(other_idxs)

            limit = max(0, int(self.session_count or 0))
            if limit:
                if limit < len(new_idxs):
                    idxs = new_idxs[:limit]
                else:
                    idxs = new_idxs + other_idxs[: max(0, limit - len(new_idxs))]
            else:
                idxs = new_idxs + other_idxs

            if self.shuffle_questions:
                self._rng.shuffle(idxs)
            return idxs

        if pick.startswith("New Qs only"):
            qids = self._new_question_qids()
            idxs = self._question_indices_by_qids(qids) if qids else list(range(len(self.questions)))
            if self.shuffle_questions:
                self._rng.shuffle(idxs)
            return idxs

        if pick.startswith("Wrong >=") or pick.startswith("Wrong \u2265"):
            try:
                thr = int(self.sp_min_wrong.value()) if hasattr(self, "sp_min_wrong") else int(self.min_wrong)
            except Exception:
                thr = 0
            thr = max(0, thr)

            qids: Set[int] = set()
            for q in self.questions:
                s = self.q_stats.get(q.qid, {}) or {}
                wrong_n = int(s.get("wrong", 0) or 0)
                if wrong_n >= thr:
                    qids.add(q.qid)

            if not qids:
                idxs = list(range(len(self.questions)))
            else:
                idxs = self._weighted_wrong_order(qids)
                weighted_mode = True

        elif pick.startswith("Wrong bank"):
            # Wrong-answer DB drill: saved bank + AI Coach + Deep Review qids.
            # Do not hide Flagged items here; this button means "show me the DB".
            qids_all = self._wrong_db_qids()
            if not qids_all:
                return list(range(len(self.questions)))
            idxs = self._weighted_wrong_order(qids_all)
            weighted_mode = True
        elif pick.startswith("Incorrect"):
            qids = set(self.last_incorrect)
            if not qids:
                return list(range(len(self.questions)))
            idxs = self._weighted_wrong_order(qids)
            weighted_mode = True
        elif pick.startswith("Flagged"):
            qids = set(self.flagged)
            if not qids:
                return list(range(len(self.questions)))
            idxs = self._question_indices_by_qids(qids)
        elif pick.startswith("Unanswered"):
            # based on current selections (if any)
            qids = {q.qid for q in self.questions if not self.selections.get(q.qid)}
            idxs = self._question_indices_by_qids(qids) if qids else list(range(len(self.questions)))
        else:
            idxs = list(range(len(self.questions)))

        if self.shuffle_questions and not weighted_mode:
            self._rng.shuffle(idxs)
        return idxs


    def _build_option_order(self) -> None:
        self.option_order = {}
        for q in self.questions:
            letters = list(q.options.keys())
            if self.shuffle_options:
                tmp = letters[:]
                self._rng.shuffle(tmp)
                self.option_order[q.qid] = tmp
            else:
                self.option_order[q.qid] = letters

    def start_session(self) -> None:
        if not self.questions:
            self.lbl_hint.setText("Load your DOCX to begin.")
            return

        # new session seed
        self._rng.seed(random.randint(1, 1_000_000_000))

        custom_ids = self._parse_custom_ids()
        self.custom_ids_active = bool(custom_ids)

        self.order = self._build_order()

        # If filters result in an empty session (common case: Wrong bank minus Flagged),
        # stop cleanly and tell the user what happened.
        if not self.order:
            QMessageBox.information(
                self,
                "No questions",
                (
                    "No questions match the selected bank/filter.\n\n"
                    "Tip: DB wrong answers comes from the saved bank plus AI Coach/Deep Review reports.\n"
                    "If it is empty here, load the DOCX that contains those qids."
                )
            )
            self._populate_nav()
            self._render_current()
            return

        # Apply session sizing (0 = ALL). If shuffle is off, still take a random sample.
        # If Custom IDs are set, we keep exactly that list.
        if (not self.custom_ids_active) and self.session_count and len(self.order) > self.session_count:
            if self.shuffle_questions:
                self.order = self.order[: self.session_count]
            else:
                self.order = self._rng.sample(self.order, self.session_count)

        self.idx = 0
        self.submitted = False
        self._session_results_recorded = False
        self._last_checked_sel = {}
        self._last_scored_sel = {}
        self.txt_feedback.setVisible(False)
        # Reset answers for a fresh session
        self.selections = {}
        self.last_incorrect = set()


        self._build_option_order()

        # timer behavior
        self._tick.stop()
        self.seconds_spent = 0

        # Reset per-session practice stats (so a 3-question session shows score out of 3)
        self.practice_correct = set()
        self.practice_attempted = set()

        # Timer behavior:
        # - Exam: countdown shown
        # - Practice: no countdown shown (label stays --:--), but we still track time spent
        if self.mode == "Exam":
            self.seconds_left = self.time_limit_min * 60
            self.lbl_timer.setText(f"\u23f1 {_fmt_mmss(self.seconds_left)}")
        else:
            self.seconds_left = None
            self.lbl_timer.setText("\u23f1 --:--")

        # Tick counts seconds_spent (and countdown if seconds_left is set)
        self._tick.start(1000)

        # NEW tracking for this session (based on historical attempts at session start)
        self.session_new_initial = set()
        for i in (self.order or []):
            qid = int(self.questions[i].qid)
            s = self.q_stats.get(qid, {}) or {}
            att = int(s.get('attempts', 0) or 0)
            if att <= 0:
                self.session_new_initial.add(qid)
        # remaining NEW are those initial NEW with no current selection
        self.session_new_remaining = set(self.session_new_initial)

        self._populate_nav()
        self._render_current()
        self._schedule_autosave()

    def start_wrong_answers(self) -> None:
        """One-click: switch bank to the wrong-answer DB and start a session."""
        if not self.questions:
            QMessageBox.warning(self, "No bank", "Load a DOCX first.")
            return

        idx = self.cb_bank.findText("Wrong bank (saved)")
        if idx >= 0:
            self.cb_bank.setCurrentIndex(idx)

        wrong_db = self._wrong_db_qids()
        if not wrong_db:
            QMessageBox.information(
                self,
                "Wrong bank empty",
                "Your wrong-answer DB is empty. Starting from all questions instead."
            )
            self.start_session()
            return

        # Honor the count spinner - same behavior as starting any other session.
        # If the spinner is 0, _build_order returns all wrongs naturally.
        self.start_session()

    def start_repeat_all_questions(self) -> None:
        """Restart the current session/profile question set without filtering by wrong/correct state."""
        if not self.questions:
            QMessageBox.warning(self, "No bank", "Load a DOCX first.")
            return

        qids: List[int] = []
        seen: Set[int] = set()
        source_idxs = self.order if getattr(self, "order", None) else range(len(self.questions))
        for i in source_idxs:
            try:
                qid = int(self.questions[i].qid)
            except Exception:
                continue
            if qid in seen:
                continue
            seen.add(qid)
            qids.append(qid)

        if not qids:
            QMessageBox.information(self, "No questions", "No questions are available to repeat.")
            return

        prev_text = self.custom_ids_text
        prev_count = self.session_count
        self._set_custom_ids_value(",".join(str(qid) for qid in qids))
        self.session_count = 0
        try:
            self.start_session()
        finally:
            self.session_count = prev_count
            self._set_custom_ids_value(prev_text)
            self.custom_ids_active = bool(self._parse_custom_ids())
            self._render_current()

    def start_round_wrong_answers(self) -> None:
        """Drill only the questions wrong in the CURRENT round (live selections),
        independent of the saved Wrong bank."""
        if not self.questions:
            QMessageBox.warning(self, "No bank", "Load a DOCX first.")
            return
        loaded = {q.qid for q in self.questions}
        round_wrong = sorted(self._live_wrong_qids() & loaded)
        if not round_wrong:
            QMessageBox.information(
                self, "Round wrong answers",
                "No wrong answers in the current round.\n\n"
                "Either you haven't answered anything wrong yet, or you already started a new "
                "session (which clears live selections). Use 'DB wrong answers' for the saved bank."
            )
            return

        # Pipe through the custom-IDs hook so _build_order respects exactly this list.
        prev_text = self.custom_ids_text
        prev_count = self.session_count
        self._set_custom_ids_value(",".join(str(q) for q in round_wrong))
        self.session_count = 0
        try:
            self.start_session()
        finally:
            self.session_count = prev_count
            # Restore the user's custom-ids box so we don't pollute it for the next session.
            self._set_custom_ids_value(prev_text)

    def show_stats(self) -> None:
        # Distribution inside wrong bank (so you can spot weak areas fast)
        bank_pattern_counts: Dict[str, int] = {}
        bank_tag_counts: Dict[str, int] = {}
        for qid in self._wrong_db_qids():
            q = self._q_by_id(qid)
            if not q:
                continue
            pid = (q.pattern_id or "").strip() or "(unlabeled)"
            bank_pattern_counts[pid] = bank_pattern_counts.get(pid, 0) + 1
            for t in (q.tags or []):
                tt = (t or "").strip()
                if not tt:
                    continue
                bank_tag_counts[tt] = bank_tag_counts.get(tt, 0) + 1

        dlg = StatsDialog(self, self.q_stats, self.p_stats, self.t_stats, bank_pattern_counts, bank_tag_counts)
        dlg.exec()

    def _live_wrong_qids(self) -> Set[int]:
        """Compute wrong qids from CURRENT selections (works mid-session in both modes).

        Counts a question as wrong if the user has made a non-empty selection
        that does not equal the correct answer. Unanswered questions do NOT
        count here - use the wrong bank or repeat offenders for those.
        """
        wrong: Set[int] = set()
        if not self.questions:
            return wrong
        # Iterate the active session order if there is one, otherwise all questions
        idxs = self.order if getattr(self, "order", None) else range(len(self.questions))
        for i in idxs:
            try:
                q = self.questions[i]
            except Exception:
                continue
            sel = set(self.selections.get(q.qid, set()))
            if not sel:
                continue
            key = set(q.effective_answer(self.overrides))
            if not key:
                continue
            if sel != key:
                wrong.add(q.qid)
        # Also union with last_incorrect (covers post-Finish state) and practice misses
        wrong |= set(getattr(self, "last_incorrect", set()))
        try:
            wrong |= (set(self.practice_attempted) - set(self.practice_correct))
        except Exception:
            pass
        return wrong

    def show_ai_coach(self) -> None:
        """Pick a wrong-answer scope and send only the next batch_size qids
        (skipping ones already analyzed for that scope). Progress persists."""
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, "AI Coach", "AI Coach is already running.")
            return

        if not self.questions:
            QMessageBox.information(self, "AI Coach", "Load a quiz first.")
            return
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()

        # Reconcile ai_coach_sent against the actual file: a qid is only truly
        # "sent" if it landed as a `### Q<N>` heading in ai_coach_reports.md.
        # Claude can drop qids on output limits even when the batch marker
        # claims it was requested. Without this, those phantom-sent qids stay
        # excluded from future batches forever.
        truly_done = self._load_master_qids()
        for _scope, _set in list(self.ai_coach_sent.items()):
            pruned = _set & truly_done
            if pruned != _set:
                self.ai_coach_sent[_scope] = pruned
        self._save_ai_coach_sent()

        loaded_qids = {q.qid for q in self.questions}

        # build the candidate sets (ALL wrong qids in scope, intersected with docx)
        current_all = sorted(self._live_wrong_qids() & loaded_qids)

        wrong_db = self._wrong_db_qids()
        bank_total = len(wrong_db)
        bank_all = sorted(wrong_db & loaded_qids)

        threshold = max(2, int(self.min_wrong))
        repeat_all = sorted([
            qid for qid, s in self.q_stats.items()
            if int(s.get("wrong", 0)) >= threshold
            and qid in loaded_qids
            and qid in wrong_db
        ])

        scope_pools = {"current": current_all, "bank": bank_all, "repeat": repeat_all}

        scopes_meta: Dict[str, Dict[str, int]] = {}
        for key, pool in scope_pools.items():
            sent = self.ai_coach_sent.get(key, set()) & set(pool)
            scopes_meta[key] = {
                "total": len(pool),
                "sent": len(sent),
                "pending": len(pool) - len(sent),
            }

        if all(m["total"] == 0 for m in scopes_meta.values()):
            QMessageBox.information(
                self, "AI Coach",
                "No wrong answers available yet.\n\n"
                "Run a session and get some wrong, or load the docx that matches your saved wrong bank."
            )
            return

        picker = AICoachPickerDialog(
            self,
            scopes=scopes_meta,
            repeat_threshold=threshold,
            wrong_bank_total=bank_total,
            batch_size_default=self.ai_coach_batch_size,
        )
        if picker.exec() != QDialog.DialogCode.Accepted or not picker.choice:
            return

        scope = picker.choice
        self.ai_coach_batch_size = picker.batch_size
        self.settings.setValue("ai_coach/batch_size", self.ai_coach_batch_size)

        if picker.reset_requested:
            self.ai_coach_sent[scope] = set()
            self._save_ai_coach_sent()
            QMessageBox.information(
                self, "AI Coach",
                f"Progress reset for scope '{scope}'. Click AI Coach again to start fresh.",
            )
            return

        pool = scope_pools[scope]
        already = self.ai_coach_sent.get(scope, set())
        master_qids = self._load_master_qids()
        master_answers = self._load_master_qid_answers()

        def _curr_ans_letter(qid: int) -> str:
            sel = sorted(self.selections.get(qid, set()))
            return ",".join(sel).upper() if sel else ""

        # A qid is pending if either:
        #   (a) it has never been analyzed (not in master, not in `already`), OR
        #   (b) the user is now picking a DIFFERENT wrong letter than every prior
        #       analyzed answer for that qid - that flip is the "no stable model"
        #       signal worth re-engaging Claude on.
        pending: List[int] = []
        flip_priors: Dict[int, List[str]] = {}
        for q in pool:
            curr = _curr_ans_letter(q)
            prior = master_answers.get(q, set())
            if prior and curr and curr not in prior:
                pending.append(q)
                flip_priors[q] = sorted(prior)
                continue
            if q in already or q in master_qids:
                continue
            pending.append(q)

        if not pending:
            in_master = len([q for q in pool if q in master_qids])
            QMessageBox.information(
                self, "AI Coach",
                f"All {len(pool)} qids in scope '{scope}' have already been analyzed "
                "with the answer you're currently picking.\n\n"
                f"({in_master} found in {self._provider_report_path('coach', provider).name}, "
                f"{len(already)} tracked in this session.)\n\n"
                "Tip: si fallas la misma con otra letra, AI Coach lo detecta como flip y "
                "vuelve a razonar. O usa 'Reset progress' para forzar.",
            )
            return

        batch_qids = pending[:picker.batch_size]
        analyzed_total = len([q for q in pool if q in already or q in master_qids])
        flip_n = sum(1 for q in batch_qids if q in flip_priors)
        scope_label = (
            f"{scope} - batch of {len(batch_qids)} "
            f"(qids {batch_qids[0]}-{batch_qids[-1]}, "
            f"{analyzed_total}/{len(pool)} already analyzed"
            + (f", {flip_n} answer-flips" if flip_n else "")
            + ")"
        )

        items: List[Dict] = []
        skipped = 0
        for qid in batch_qids:
            try:
                q = self._q_by_id(qid)
            except KeyError:
                skipped += 1
                continue
            wrong_n = int(self.q_stats.get(qid, {}).get("wrong", 0) or 0)
            sel = sorted(self.selections.get(qid, set()))
            your = ",".join(sel) if sel else "(no answer recorded)"
            item: Dict = {
                "qid": q.qid,
                "stem": q.stem,
                "options": dict(q.options),
                "correct_answer": q.effective_answer(self.overrides),
                "your_answer": your,
                "explanation": (self.my_expl.get(q.qid) or q.explanation or ""),
                "pattern_id": q.pattern_id or "",
                "tags": list(q.tags or []),
                "lifetime_wrong_count": wrong_n,
            }
            if qid in flip_priors:
                item["prior_wrong_answers"] = flip_priors[qid]
            items.append(item)

        if not items:
            QMessageBox.information(self, "AI Coach", "Nothing to send (all qids skipped).")
            return

        sent_n = len(items)
        sent_qids = {it["qid"] for it in items}

        # Preflight in main thread: catch missing SDK / wrong python BEFORE we
        # spawn a QThread (where errors are easy to lose behind the modal dialog).
        try:
            import quiz_ai_coach as _qac
            _qac._logger.info("preflight ok from GUI (python=%s)", sys.executable)
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, "AI Coach - cannot start",
                f"Failed to import quiz_ai_coach / claude_agent_sdk.\n\n"
                f"Python: {sys.executable}\n\n"
                f"{type(_e).__name__}: {_e}\n\n"
                f"Fix: run\n  \"{sys.executable}\" -m pip install claude-agent-sdk\n\n"
                f"{_tb.format_exc()}",
            )
            return

        # Scale timeout with batch size: ~45s/question + 60s baseline, capped at 15min.
        dyn_timeout = min(900.0, 60.0 + 45.0 * sent_n)

        progress = QProgressDialog(
            f"Asking {backend_label} to analyze: {scope_label}...\n"
            f"(timeout {int(dyn_timeout)}s - first chars usually arrive in 5-15s; "
            f"big batches can take a few minutes)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("AI Coach")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_ai_coach.setEnabled(False)
        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "scope": scope,
            "ai_provider": provider,
            "pool": list(pool),
            "sent_n": sent_n,
            "sent_qids": set(sent_qids),
            "scope_label": scope_label,
            "ai_backend_label": backend_label,
            "tick": 0,
            "last_msg": "waiting for first chunk...",
            "cancelled": False,
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        self._ai_coach_worker = AICoachWorker(items, model=ai_model, provider=provider, timeout_sec=dyn_timeout)
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._ai_coach_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_thread.start()

    def show_top_offenders(self) -> None:
        """List loaded-bank questions ranked by lifetime wrong count. Click Q### to jump."""
        rows = []
        loaded = {q.qid: q for q in (self.questions or [])}
        qids = sorted(loaded.keys()) if loaded else sorted((self.q_stats or {}).keys())
        for qid in qids:
            s = (self.q_stats or {}).get(qid, {}) or {}
            w = int(s.get("wrong", 0) or 0)
            a = int(s.get("attempts", 0) or 0)
            c = int(s.get("correct", 0) or 0)
            rate = (w / a * 100.0) if a else 0.0
            rows.append((qid, w, a, c, rate))
        if not rows:
            QMessageBox.information(self, "Top Offenders",
                "No hay preguntas cargadas ni estadisticas guardadas todavia.")
            return
        rows.sort(key=lambda r: (-r[1], -r[4], -r[2], r[0]))
        scope = "preguntas del banco cargado" if loaded else "preguntas con estadisticas guardadas"
        lines = [f"# \U0001f525 Top Offenders - {len(rows)} {scope}\n",
                 "Ordenado por *lifetime wrong count* (luego wrong-rate). Incluye preguntas con 0 fallos al final. Click en `Q###` para saltar.\n",
                 "| # | QID | Wrong | Attempts | Correct | Rate | Tema |",
                 "|---|---|---|---|---|---|---|"]
        for i, (qid, w, a, c, rate) in enumerate(rows, 1):
            q = loaded.get(qid)
            stem = (q.stem if q else "").strip().replace("\n", " ").replace("|", "\\|")
            stem = re.sub(r"\s+", " ", stem)
            if len(stem) > 90:
                stem = stem[:89] + "..."
            lines.append(f"| {i} | Q{qid} | **{w}** | {a} | {c} | {rate:.0f}% | {stem} |")
        report_md = "\n".join(lines)

        dlg = AICoachDialog(
            self, report_md=report_md, num_questions=len(rows),
            model="(local stats)", master_path=None,
            header_text=f"Top Offenders - {len(rows)} {scope}",
        )
        dlg.setWindowTitle("\U0001f525 Top Offenders")
        try:
            self._open_coach_dialogs.append(dlg)
        except Exception:
            pass
        dlg.show()

    def start_due_today(self) -> None:
        """Start a session of qids whose SRS next_review <= today."""
        if not self.questions:
            QMessageBox.warning(self, "No bank", "Load a DOCX first.")
            return
        srs = _load_json(_here("srs_state.json"), {}) or {}
        today = _today_iso()
        bank_qids = {q.qid for q in self.questions}
        due: List[int] = []
        for k, v in srs.items():
            try:
                qid = int(k)
            except Exception:
                continue
            if qid not in bank_qids:
                continue
            nxt = str((v or {}).get("next_review", "") or "")
            if nxt and nxt <= today:
                due.append(qid)
        # Include qids that have NEVER been seen (no entry in srs) - they're "due" by definition.
        never_seen = sorted(bank_qids - {int(k) for k in srs.keys() if str(k).isdigit()})
        candidates = sorted(set(due) | set(never_seen))
        if not candidates:
            QMessageBox.information(self, "Due Today",
                "No hay preguntas vencidas. Vuelve manana o usa otro scope.")
            return
        msg = (f"{len(due)} preguntas vencidas (SRS) + {len(never_seen)} nunca vistas "
               f"= {len(candidates)} en cola.\n\nIniciar sesion con todas?")
        if QMessageBox.question(self, "\U0001f5d3\ufe0f Due Today", msg) != QMessageBox.Yes:
            return
        idx = self.cb_bank.findText("Wrong bank (saved)")
        if idx >= 0:
            self.cb_bank.setCurrentIndex(idx)
        self._due_today_override = set(candidates)
        prev_count = self.session_count
        self.session_count = 0
        try:
            self.start_session()
        finally:
            self.session_count = prev_count

    def show_concept_mastery(self) -> None:
        """Render concept_mastery.json as a sortable markdown table."""
        cm = _load_json(_here("concept_mastery.json"), {}) or {}
        if not cm:
            QMessageBox.information(self, "Mastery",
                "concept_mastery.json esta vacio. Responde algunas preguntas primero.")
            return
        rows = []
        for name, c in cm.items():
            r = int(c.get("right", 0)); w = int(c.get("wrong", 0)); a = int(c.get("attempts", 0))
            rate = (r / a * 100.0) if a else 0.0
            rows.append((name, r, w, a, rate, c.get("last_seen", ""), len(c.get("qids") or [])))
        rows.sort(key=lambda x: (x[4], -x[3], x[0]))  # worst mastery first, then most-attempted
        lines = ["# \U0001f3af Concept Mastery\n",
                 "Ordenado por mastery (peor primero). Concepto detectado via CONCEPT_CATALOG.\n",
                 "| # | Concept | Mastery | Right | Wrong | Attempts | Qids | Last seen |",
                 "|---|---|---|---|---|---|---|---|"]
        for i, (name, r, w, a, rate, last, nq) in enumerate(rows, 1):
            bar = "\U0001f7e2" if rate >= 80 else ("\U0001f7e1" if rate >= 60 else "\U0001f534")
            lines.append(f"| {i} | {name} | {bar} {rate:.0f}% | {r} | **{w}** | {a} | {nq} | {last} |")
        report_md = "\n".join(lines)
        dlg = AICoachDialog(
            self, report_md=report_md, num_questions=len(rows),
            model="(local stats)", master_path=_here("concept_mastery.json"),
            header_text=f"Mastery - {len(rows)} concepts tracked",
        )
        dlg.setWindowTitle("\U0001f3af Concept Mastery")
        try:
            self._open_coach_dialogs.append(dlg)
        except Exception:
            pass
        dlg.show()

    def show_db_report(self) -> None:
        """Render a consolidated local report of ai_coach_reports.md.

        No Claude call. Shows summary stats (totals, scopes, date range, batches)
        plus the full deduped DB content in the same dialog you already use for
        AI Coach reports.
        """
        master_path = self._provider_report_path("coach")
        if not master_path.exists():
            QMessageBox.information(
                self, "DB Report",
                f"{master_path.name} does not exist yet.\n\n"
                "Run AI Coach at least once to populate the DB.",
            )
            return
        try:
            text = master_path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            QMessageBox.critical(self, "DB Report", f"Could not read {master_path}:\n{e}")
            return
        if not text.strip():
            QMessageBox.information(self, "DB Report", "DB file is empty.")
            return

        all_qids: Set[int] = set()
        scopes: Dict[str, int] = {}
        timestamps: List[str] = []
        for m in re.finditer(
            r"<!--\s*AI_COACH_BATCH\s+qids=\[([^\]]*)\]\s+scope=(\S+)\s+ts=([0-9\-:\s]+?)\s*-->",
            text,
        ):
            for tok in m.group(1).split(","):
                tok = tok.strip()
                if tok.lstrip("-").isdigit():
                    all_qids.add(int(tok))
            scopes[m.group(2)] = scopes.get(m.group(2), 0) + 1
            timestamps.append(m.group(3).strip())
        batch_count = len(timestamps)
        ts_min = min(timestamps) if timestamps else "-"
        ts_max = max(timestamps) if timestamps else "-"
        scopes_line = ", ".join(f"{k}: {v}" for k, v in sorted(scopes.items())) or "-"

        header = (
            f"# AI Coach DB - consolidated report\n\n"
            f"- **File**: `{master_path}`\n"
            f"- **Size**: {len(text):,} chars\n"
            f"- **Batches**: {batch_count}\n"
            f"- **Unique qids analyzed**: {len(all_qids)}\n"
            f"- **Scopes**: {scopes_line}\n"
            f"- **Date range**: {ts_min} -> {ts_max}\n"
            f"- **Qids**: {', '.join(str(q) for q in sorted(all_qids)) or '-'}\n\n"
            f"---\n\n"
        )

        # Backfill recaps locally from the loaded docx (no tokens) for sections
        # written before the recap was added to the AI Coach prompt.
        text = self._inject_recaps_inline(text)

        dlg = AICoachDialog(
            self, header + text, num_questions=len(all_qids),
            model="local \u00b7 DB report (no Claude call)",
            master_path=master_path,
        )
        dlg.setWindowTitle("AI Coach DB - consolidated report")
        self._show_coach_dialog_nonmodal(dlg)

    def show_deep_review_md(self) -> None:
        """Open the entire deep_review_reports.md (no Claude call)."""
        deep_path = self._provider_report_path("deep")
        if not deep_path.exists():
            QMessageBox.information(
                self, "Show Deep Review",
                f"{deep_path.name} does not exist yet.\n\n"
                "Run Deep Review at least once to populate the file.",
            )
            return
        try:
            text = deep_path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            QMessageBox.critical(self, "Show Deep Review", f"Could not read {deep_path}:\n{e}")
            return
        if not text.strip():
            QMessageBox.information(self, "Show Deep Review", "Deep Review file is empty.")
            return

        all_qids: Set[int] = set()
        scopes: Dict[str, int] = {}
        timestamps: List[str] = []
        for m in re.finditer(
            r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)\s+scope=(\S+)\s+ts=([0-9\-:\s]+?)\s+status=",
            text,
        ):
            all_qids.add(int(m.group(1)))
            scopes[m.group(2)] = scopes.get(m.group(2), 0) + 1
            timestamps.append(m.group(3).strip())
        ts_min = min(timestamps) if timestamps else "-"
        ts_max = max(timestamps) if timestamps else "-"
        scopes_line = ", ".join(f"{k}: {v}" for k, v in sorted(scopes.items())) or "-"

        header = (
            f"# Deep Review - full archive\n\n"
            f"- **File**: `{deep_path}`\n"
            f"- **Size**: {len(text):,} chars\n"
            f"- **Unique qids deep-reviewed**: {len(all_qids)}\n"
            f"- **Scopes**: {scopes_line}\n"
            f"- **Date range**: {ts_min} -> {ts_max}\n"
            f"- **Qids**: {', '.join(str(q) for q in sorted(all_qids)) or '-'}\n\n"
            f"---\n\n"
        )

        dlg = AICoachDialog(
            self, header + text, num_questions=len(all_qids),
            model="local \u00b7 Deep Review archive (no Claude call)",
            master_path=deep_path,
        )
        dlg.setWindowTitle("Deep Review - full archive")
        self._show_coach_dialog_nonmodal(dlg)

    def show_nuclear_md(self) -> None:
        """Open the entire nuclear_reports.md (no Claude call)."""
        nuclear_path = self._provider_report_path("nuclear")
        if not nuclear_path.exists():
            QMessageBox.information(
                self, "Show Nuclear",
                f"{nuclear_path.name} does not exist yet.\n\n"
                "Run Nuclear Review at least once to populate the file.",
            )
            return
        try:
            text = nuclear_path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            QMessageBox.critical(self, "Show Nuclear", f"Could not read {nuclear_path}:\n{e}")
            return
        if not text.strip():
            QMessageBox.information(self, "Show Nuclear", "Nuclear file is empty.")
            return

        all_qids: Set[int] = set()
        timestamps: List[str] = []
        run_count = 0
        for m in re.finditer(
            r"<!--\s*NUCLEAR_REVIEW\s+qid=(\d+)[^>]*ts=([0-9\-:\s]+?)\s*-->",
            text,
        ):
            all_qids.add(int(m.group(1)))
            timestamps.append(m.group(2).strip())
            run_count += 1
        ts_min = min(timestamps) if timestamps else "-"
        ts_max = max(timestamps) if timestamps else "-"

        header = (
            f"# Nuclear Review - full archive\n\n"
            f"- **File**: `{nuclear_path}`\n"
            f"- **Size**: {len(text):,} chars\n"
            f"- **Total runs**: {run_count}\n"
            f"- **Unique qids**: {len(all_qids)}\n"
            f"- **Date range**: {ts_min} -> {ts_max}\n"
            f"- **Qids**: {', '.join(str(q) for q in sorted(all_qids)) or '-'}\n\n"
            f"---\n\n"
        )

        dlg = AICoachDialog(
            self, header + text, num_questions=len(all_qids),
            model="local \u00b7 Nuclear archive (no Claude call)",
            master_path=nuclear_path,
        )
        dlg.setWindowTitle("Nuclear Review - full archive")
        self._show_coach_dialog_nonmodal(dlg)

    def show_deep_review_md_current_row(self) -> None:
        """Open deep_review_reports.md filtered to the qids the user got wrong in the last submitted row."""
        wrong_qids: Set[int] = set(getattr(self, "last_incorrect", set()) or set())
        self._show_deep_review_filtered(
            wrong_qids,
            label="current row wrongs",
            empty_msg="No wrongs recorded for the current row yet.\n\nSubmit a row first.",
        )

    def show_deep_review_md_current_round(self) -> None:
        """Open deep_review_reports.md filtered to the qids in the current quiz round (self.order)."""
        round_qids: Set[int] = set()
        try:
            if self.questions and self.order:
                round_qids = {int(self.questions[i].qid) for i in self.order}
        except Exception:
            round_qids = set()
        self._show_deep_review_filtered(
            round_qids,
            label="current round",
            empty_msg="No active quiz round.\n\nStart a session first.",
        )

    def _show_deep_review_filtered(self, qids: Set[int], *, label: str, empty_msg: str) -> None:
        if not qids:
            QMessageBox.information(self, f"Show Deep Review ({label})", empty_msg)
            return
        deep_path = self._provider_report_path("deep")
        if not deep_path.exists():
            QMessageBox.information(
                self, f"Show Deep Review ({label})",
                f"{deep_path.name} does not exist yet.",
            )
            return
        try:
            text = deep_path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            QMessageBox.critical(self, f"Show Deep Review ({label})", f"Could not read {deep_path}:\n{e}")
            return

        marker_re = re.compile(r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)[^>]*-->")
        matches = list(marker_re.finditer(text))
        sections: List[str] = []
        found_qids: Set[int] = set()
        for i, m in enumerate(matches):
            qid = int(m.group(1))
            if qid in qids:
                start = m.start()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                sections.append(text[start:end].rstrip())
                found_qids.add(qid)

        missing = sorted(qids - found_qids)
        header = (
            f"# Deep Review - {label}\n\n"
            f"- **File**: `{deep_path}`\n"
            f"- **Qids in {label}**: {len(qids)} ({', '.join(str(q) for q in sorted(qids))})\n"
            f"- **Found in archive**: {len(found_qids)}\n"
        )
        if missing:
            header += f"- **Missing (run Deep Review on these)**: {', '.join(str(q) for q in missing)}\n"
        header += "\n---\n\n"

        body = "\n\n".join(sections) if sections else "_No matching deep reviews found in archive._"
        dlg = AICoachDialog(
            self, header + body, num_questions=len(found_qids),
            model="local \u00b7 Deep Review filtered (no Claude call)",
            master_path=deep_path,
        )
        dlg.setWindowTitle(f"Deep Review - {label}")
        self._show_coach_dialog_nonmodal(dlg)

    def _meta_history_path(self, provider: Optional[str] = None) -> Path:
        return self._provider_report_path("meta_history", provider)

    def _load_meta_history(self, provider: Optional[str] = None) -> List[Dict]:
        p = self._meta_history_path(provider)
        if not p.exists():
            return []
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else []
        except Exception:
            return []

    def _append_meta_history(self, stats_payload: Dict, provider: Optional[str] = None) -> None:
        history = self._load_meta_history(provider)
        snapshot = {
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "totals": stats_payload.get("totals", {}),
            "hot_zone_qids": [h["qid"] for h in stats_payload.get("hot_zone", [])],
            "top_reincident_qids": [r["qid"] for r in stats_payload.get("top_reincident", [])],
            "pending_qids": list(stats_payload.get("pending_qids", [])),
        }
        history.append(snapshot)
        try:
            self._meta_history_path(provider).write_text(
                json.dumps(history, indent=2, ensure_ascii=False), encoding="utf-8",
            )
        except Exception:
            pass

    def _meta_progress_section(self, stats_payload: Dict, provider: Optional[str] = None) -> str:
        """Compare current snapshot against previous one and produce a verdict.

        The verdict is the first thing the user sees in the report so the
        'sigues pendejo vs veo mejora' question is answered before they
        scroll into the analysis.
        """
        history = self._load_meta_history(provider)
        if not history:
            return (
                "## Progreso vs corrida anterior\n\n"
                "_Primera corrida de Meta-Coach. Proxima vez vas a ver aca si "
                "mejoraste o sigues pendejo._\n\n---\n\n"
            )
        prev = history[-1]
        prev_hot = set(prev.get("hot_zone_qids", []))
        prev_unique = int((prev.get("totals") or {}).get("unique_qids_in_master", 0) or 0)
        prev_pending = int((prev.get("totals") or {}).get("pending_in_bank", 0) or 0)

        curr_hot = {h["qid"] for h in stats_payload.get("hot_zone", [])}
        curr_unique = int((stats_payload.get("totals") or {}).get("unique_qids_in_master", 0) or 0)
        curr_pending = int((stats_payload.get("totals") or {}).get("pending_in_bank", 0) or 0)

        fixed = sorted(prev_hot - curr_hot)      # estaban hot, ya no
        new_hot = sorted(curr_hot - prev_hot)    # nuevos en hot zone
        still_hot = sorted(prev_hot & curr_hot)  # siguen ahi (lo mas grave)
        delta_hot = len(curr_hot) - len(prev_hot)
        delta_unique = curr_unique - prev_unique

        if len(curr_hot) == 0 and len(prev_hot) > 0:
            verdict = "\U0001f7e2 **VEO MEJORA** - limpiaste toda la hot zone."
        elif delta_hot < 0 and len(still_hot) <= 1:
            verdict = f"\U0001f7e2 **VEO MEJORA** - hot zone bajo de {len(prev_hot)} a {len(curr_hot)}."
        elif delta_hot < 0:
            verdict = (
                f"\U0001f7e1 **MEJORA PARCIAL** - hot zone bajo de {len(prev_hot)} a {len(curr_hot)} "
                f"pero {len(still_hot)} qids reinciden hace 2 corridas."
            )
        elif delta_hot == 0 and len(still_hot) == len(curr_hot) and len(curr_hot) > 0:
            verdict = (
                f"\U0001f534 **SIGUES PENDEJO** - los mismos {len(curr_hot)} qids siguen en hot zone "
                f"sin moverse desde la corrida anterior."
            )
        elif delta_hot > 0:
            verdict = (
                f"\U0001f534 **SIGUES PENDEJO** - hot zone crecio de {len(prev_hot)} a {len(curr_hot)} "
                f"({len(new_hot)} qids nuevos haciendo guess)."
            )
        else:
            verdict = "\U0001f7e1 **ESTABLE** - sin cambios significativos en hot zone."

        lines = [
            "## Progreso vs corrida anterior",
            "",
            verdict,
            "",
            f"- **Corrida anterior**: {prev.get('ts', '?')}",
            f"- **Hot zone**: {len(prev_hot)} -> {len(curr_hot)} (delta {delta_hot:+d})",
            f"- **Qids unicos en master**: {prev_unique} -> {curr_unique} (delta {delta_unique:+d})",
            f"- **Pendientes en bank**: {prev_pending} -> {curr_pending}",
        ]
        if fixed:
            lines.append(f"- \u2705 **Limpiados de hot zone**: {', '.join(f'Q{q}' for q in fixed)}")
        if still_hot:
            lines.append(
                f"- \u26a0\ufe0f **Reinciden (2+ corridas en hot zone)**: "
                f"{', '.join(f'Q{q}' for q in still_hot)} <- prioridad maxima"
            )
        if new_hot:
            lines.append(f"- \U0001f195 **Nuevos en hot zone**: {', '.join(f'Q{q}' for q in new_hot)}")
        lines.extend(["", "---", "", ""])
        return "\n".join(lines)

    def _build_meta_stats_payload(self, reports_md: str) -> Dict:
        """Build the JSON sidecar for Meta-Coach.

        - hot_zone: qids that appear in 2+ batches with DIFFERENT user answers
          (the user keeps guessing, doesn't have a stable mental model)
        - top_reincident: top qids by lifetime wrong count
        - totals: counts so the model knows what's covered vs pending
        """
        loaded_qids = {q.qid for q in self.questions} if self.questions else set()
        master_qids = self._load_master_qids()
        bank = self._wrong_db_qids()
        bank_in_loaded = bank & loaded_qids if loaded_qids else bank
        pending = sorted(bank_in_loaded - master_qids)

        # Parse per-qid user answers across batches from the reports markdown.
        # Each batch's report has sections like "### Q123 - topic" then a line
        # "- **Your answer:** X - ..." (X is a letter or comma list).
        per_qid_answers: Dict[int, List[str]] = {}
        # Split master into batches by the comment marker so a qid analyzed in
        # batch A and batch B counts as 2 occurrences even with same letter.
        batches = re.split(r"<!--\s*AI_COACH_BATCH[^>]*-->", reports_md)
        for batch_text in batches:
            seen_in_batch: Set[int] = set()
            for m in re.finditer(
                r"###\s*Q(\d+)\b[^\n]*\n(?:[^\n]*\n){0,3}?[^\n]*\*\*Your answer:\*\*\s*([A-Za-z](?:\s*,\s*[A-Za-z])*)",
                batch_text,
            ):
                qid = int(m.group(1))
                if qid in seen_in_batch:
                    continue
                seen_in_batch.add(qid)
                ans = re.sub(r"\s+", "", m.group(2)).upper()
                per_qid_answers.setdefault(qid, []).append(ans)

        hot_zone = []
        for qid, answers in per_qid_answers.items():
            if len(answers) >= 2 and len(set(answers)) >= 2:
                hot_zone.append({
                    "qid": qid,
                    "occurrences": len(answers),
                    "answers_given": answers,
                    "lifetime_wrong": int(self.q_stats.get(qid, {}).get("wrong", 0) or 0),
                })
        hot_zone.sort(key=lambda x: (-x["occurrences"], -x["lifetime_wrong"], x["qid"]))

        top_reincident = sorted(
            [
                {
                    "qid": qid,
                    "lifetime_wrong": int(s.get("wrong", 0) or 0),
                    "attempts": int(s.get("attempts", 0) or 0),
                }
                for qid, s in self.q_stats.items()
                if int(s.get("wrong", 0) or 0) >= 2
            ],
            key=lambda x: (-x["lifetime_wrong"], -x["attempts"], x["qid"]),
        )[:15]

        return {
            "totals": {
                "unique_qids_in_master": len(master_qids),
                "wrong_bank_total": len(bank),
                "wrong_bank_in_loaded_docx": len(bank_in_loaded),
                "pending_in_bank": len(pending),
                "batches_in_master": reports_md.count("<!-- AI_COACH_BATCH"),
            },
            "hot_zone": hot_zone[:10],
            "top_reincident": top_reincident,
            "pending_qids": pending[:30],
        }

    def show_deep_review(self) -> None:
        """Per-question deep c\u00e1tedra. Standalone (no AI Coach prerequisite) but
        uses prior AI Coach / Pre-Brief reports as `prior_report` context when present."""
        self._run_deep_or_prebrief(mode="deep")

    def show_pre_brief(self) -> None:
        """Cheap context dossier per qid. Output appended to ai_coach_reports.md
        so Deep Review picks it up later as `prior_report`."""
        self._run_deep_or_prebrief(mode="prebrief")

    def _concept_dossier_dir(self) -> Path:
        return self._provider_dossier_dir()

    def _concept_dossier_title(self, path: Path) -> str:
        try:
            text = path.read_text(encoding="utf-8", errors="replace")[:5000]
        except Exception:
            text = ""

        m = re.search(r"<!--\s*concept:\s*(.+?)(?:\s*(?:\u00b7|\u00c2\u00b7|\|)\s*|\s*-->)", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()

        m = re.search(r"^#\s+(.+?)\s*$", text, flags=re.MULTILINE)
        if m:
            return m.group(1).strip()

        return path.stem.replace("-", " ").title()

    def _concept_dossier_files(self) -> List[Path]:
        dossier_dir = self._concept_dossier_dir()
        if not dossier_dir.exists():
            return []
        files = [p for p in dossier_dir.glob("*.md") if p.is_file()]
        return sorted(files, key=lambda p: self._concept_dossier_title(p).lower())

    def _rebuild_concept_dossier_menu(self) -> None:
        menu = getattr(self, "_dossier_menu", None)
        if menu is None:
            return

        menu.clear()
        files = self._concept_dossier_files()
        if not files:
            action = menu.addAction("No concept dossiers yet")
            action.setEnabled(False)
            menu.addSeparator()
            menu.addAction("Build dossiers first...", self.build_concept_dossiers)
            return

        for path in files:
            title = self._concept_dossier_title(path)
            action = menu.addAction(title)
            action.setToolTip(str(path))
            action.triggered.connect(lambda _checked=False, p=path: self.show_concept_dossier(p))

        menu.addSeparator()
        menu.addAction("Open concept_dossiers folder", self.open_concept_dossier_folder)

    def open_concept_dossier_folder(self) -> None:
        dossier_dir = self._concept_dossier_dir()
        dossier_dir.mkdir(exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(dossier_dir)))

    def show_concept_dossier(self, path: Path) -> None:
        path = Path(path)
        if not path.exists():
            QMessageBox.warning(self, "Concept Dossier", f"No existe:\n{path}")
            return
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            QMessageBox.critical(self, "Concept Dossier", f"No pude leer {path}:\n{e}")
            return

        title = self._concept_dossier_title(path)
        dlg = AICoachDialog(
            self,
            report_md=text,
            num_questions=0,
            model="local \u00b7 Concept dossier (no Claude call)",
            master_path=path,
            header_text=f"Concept Dossier - {title}",
        )
        dlg.setWindowTitle(f"Concept Dossier - {title}")
        dlg.exec()

    def build_concept_dossiers(self) -> None:
        """Parse deep_review_reports.md (and ai_coach_reports.md), group sections
        by AWS concept, and consolidate each group via one Claude call into a
        canonical dossier saved under concept_dossiers/<slug>.md.

        This is the systemic improvement: future Deep / Nuclear Reviews retrieve
        the dossier instead of re-deriving theory from scratch."""
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, "Concept Dossiers", "A coach call is already running.")
            return
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()

        deep_path = self._provider_report_path("deep")
        master_path = self._provider_report_path("coach")
        sources: List[tuple] = []  # (qid, topic, body, source_file)

        for path in (deep_path, master_path):
            if not path.exists():
                continue
            try:
                text = path.read_text(encoding="utf-8", errors="replace")
            except Exception:
                continue
            sections = re.split(r"(?=^###\s+Q\d+)", text, flags=re.MULTILINE)
            for sec in sections:
                m = re.match(r"^###\s+Q(\d+)\b[^\n]*?(?:(?:\u2014|-)\s*(.+?))?\n", sec)
                if not m:
                    continue
                try:
                    qid = int(m.group(1))
                except ValueError:
                    continue
                topic = (m.group(2) or "").strip()
                body = sec[m.end():].strip()
                if not body:
                    continue
                sources.append((qid, topic, body, path.name))

        if not sources:
            QMessageBox.information(
                self, "Concept Dossiers",
                f"No source material found in {deep_path.name} or {master_path.name}.\n\n"
                "Run Deep Review on some qids first."
            )
            return

        # Group by concept. A section can contribute to multiple concepts.
        concept_groups: Dict[str, List[Dict[str, Any]]] = {}
        for qid, topic, body, src in sources:
            haystack = f"{topic}\n{body}"
            concepts = _detect_concepts(haystack)
            for c in concepts:
                concept_groups.setdefault(c, []).append({
                    "qid": qid,
                    "topic": topic or "(no topic)",
                    "body": body[:4000],
                    "source": src,
                })

        if not concept_groups:
            QMessageBox.information(
                self, "Concept Dossiers",
                f"Parsed {len(sources)} sections but matched ZERO concepts in the catalog.\n\n"
                "Concept catalog might need to be extended. Check CONCEPT_CATALOG in the code."
            )
            return

        # Existing dossiers - let user pick: refresh all, or only missing.
        dossier_dir = self._provider_dossier_dir(provider)
        existing = set()
        if dossier_dir.exists():
            for p in dossier_dir.glob("*.md"):
                existing.add(p.stem)

        all_concepts = sorted(concept_groups.keys(), key=lambda c: -len(concept_groups[c]))
        missing_concepts = [c for c in all_concepts if _slug_for_concept(c) not in existing]

        if not missing_concepts and existing:
            ans = QMessageBox.question(
                self, "Concept Dossiers",
                f"All {len(all_concepts)} detected concepts already have dossiers in "
                f"concept_dossiers/.\n\n"
                "Rebuild ALL anyway? Existing files are overwritten by concept slug; "
                "no duplicate dossier files are created.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            )
            if ans != QMessageBox.StandardButton.Yes:
                return
            target_concepts = all_concepts
        elif missing_concepts and existing:
            ans = QMessageBox.question(
                self, "Concept Dossiers",
                f"Found {len(all_concepts)} concepts. {len(existing)} already cached, "
                f"{len(missing_concepts)} missing.\n\n"
                "Yes = build only missing (no duplicates).  "
                "No = rebuild ALL and overwrite the same concept files.  Cancel = abort.",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
                | QMessageBox.StandardButton.Cancel,
            )
            if ans == QMessageBox.StandardButton.Cancel:
                return
            target_concepts = missing_concepts if ans == QMessageBox.StandardButton.Yes else all_concepts
        else:
            target_concepts = all_concepts

        # Confirm cost.
        ans = QMessageBox.question(
            self, "Concept Dossiers",
            f"Will run {len(target_concepts)} {backend_label} calls (one per concept), "
            f"up to {DOSSIER_MAX_WORKERS} in parallel.\n\n"
            "Each concept maps to one stable file in concept_dossiers/<slug>.md; "
            "rebuilds overwrite that file instead of creating duplicates.\n\n"
            f"Each concept uses the best {DOSSIER_MAX_EXCERPTS_PER_CONCEPT} unique-qid excerpts "
            f"max, capped at {DOSSIER_MAX_CHARS_PER_EXCERPT} chars each, so the calls stay focused.\n\n"
            "Concepts:\n  - "
            + "\n  - ".join(target_concepts[:15])
            + (f"\n  ... + {len(target_concepts)-15} more" if len(target_concepts) > 15 else "")
            + "\n\nProceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        groups_payload = [
            {
                "concept": c,
                "excerpts": _pack_dossier_excerpts(concept_groups[c]),
                "raw_excerpt_count": len(concept_groups[c]),
            }
            for c in target_concepts
        ]

        try:
            import quiz_ai_coach as _qac
            _qac._logger.info("dossier preflight ok (python=%s)", sys.executable)
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, "Concept Dossiers - cannot start",
                f"Failed to import quiz_ai_coach.\n\n{type(_e).__name__}: {_e}\n\n{_tb.format_exc()}",
            )
            return

        dossier_dir.mkdir(exist_ok=True)

        progress = QProgressDialog(
            f"Building {len(target_concepts)} concept dossiers...\n"
            f"({DOSSIER_MAX_WORKERS} parallel workers, cached files save as each finishes)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("Concept Dossiers")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_dossiers.setEnabled(False)
        self._dossier_progress = progress
        self._dossier_dir = dossier_dir
        self._dossier_count = 0
        self._dossier_total = len(target_concepts)

        self._dossier_thread = QThread(self)
        self._dossier_worker = ConceptDossierWorker(
            groups_payload,
            model=ai_model,
            provider=provider,
            timeout_sec_per_concept=360.0,
            max_workers=DOSSIER_MAX_WORKERS,
        )
        self._dossier_worker.moveToThread(self._dossier_thread)
        self._dossier_thread.started.connect(self._dossier_worker.run)
        progress.canceled.connect(self._cancel_dossier_build)
        self._dossier_worker.progress_text.connect(
            self._dossier_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._dossier_worker.one_done.connect(
            self._dossier_one_done,
            Qt.ConnectionType.QueuedConnection,
        )
        self._dossier_worker.finished.connect(
            self._dossier_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._dossier_thread.start()

    def _dossier_progress_update(self, msg: str) -> None:
        if getattr(self, "_dossier_progress", None):
            try:
                self._dossier_progress.setLabelText(
                    f"[{self._dossier_count}/{self._dossier_total}] {msg}"
                )
            except Exception:
                pass

    def _dossier_one_done(self, concept: str, slug: str, dossier_md: str) -> None:
        # Save incrementally so a crash mid-batch doesn't lose work.
        try:
            out_path = self._dossier_dir / f"{slug}.md"
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            out_path.write_text(
                f"<!-- concept: {concept} \u00b7 generated: {ts} -->\n\n{dossier_md}\n",
                encoding="utf-8",
            )
            self._dossier_count += 1
        except Exception as _ex:
            print(f"[dossier save error] {concept}: {_ex}", file=sys.stderr)

    def _cancel_dossier_build(self) -> None:
        # Workers don't currently support mid-flight cancel of in-progress Claude
        # call, but flagging stops queueing further concepts.
        if getattr(self, "_dossier_worker", None):
            try:
                self._dossier_worker.cancel_requested = True
                self._dossier_worker.groups = []
            except Exception:
                pass

    def _dossier_finished(self, results: List[tuple], error: str) -> None:
        progress = getattr(self, "_dossier_progress", None)
        thread = getattr(self, "_dossier_thread", None)
        worker = getattr(self, "_dossier_worker", None)
        self._dossier_progress = None
        self._dossier_thread = None
        self._dossier_worker = None
        if progress:
            progress.close()
            progress.deleteLater()
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        try:
            self.btn_dossiers.setEnabled(True)
        except Exception:
            pass

        n_built = len(results)
        if error:
            QMessageBox.warning(
                self, "Concept Dossiers",
                f"Built {n_built} dossiers before error:\n\n{error[:1500]}\n\n"
                f"Saved to: {self._dossier_dir}",
            )
        else:
            QMessageBox.information(
                self, "Concept Dossiers",
                f"Built {n_built} concept dossiers.\n\nSaved to: {self._dossier_dir}",
            )

    def show_nuclear_review(self) -> None:
        """Ultradios mode for ONE critical question. Spawns 3 parallel sub-agents
        (boundary, patterns, distractors) + a synthesizer. Loads matching concept
        dossiers as backbone. Saves to nuclear_reports.md."""
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, "Nuclear Review", "A coach call is already running.")
            return
        if not self.questions:
            QMessageBox.information(self, "Nuclear Review", "Load a quiz first.")
            return
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()

        # Pick qid: default to most-failed live wrong, otherwise prompt.
        loaded_qids = {q.qid for q in self.questions}
        live_wrong = sorted(self._live_wrong_qids() & loaded_qids)
        wrong_counts = {qid: int(s.get("wrong", 0) or 0) for qid, s in self.q_stats.items()}
        live_wrong_sorted = sorted(live_wrong, key=lambda q: (-wrong_counts.get(q, 0), q))
        default_qid = live_wrong_sorted[0] if live_wrong_sorted else (
            sorted(self.q_stats.keys(), key=lambda q: (-wrong_counts.get(q, 0), q))[:1] or [None]
        )[0]

        from PySide6.QtWidgets import QInputDialog
        qid_str, ok = QInputDialog.getText(
            self, "\u2622 Nuclear Review",
            "Pick ONE qid for the ultradios review.\n"
            "Spawns 3 parallel sub-agents + synthesizer (~3-5 min total).\n\n"
            f"Live wrong (sorted by lifetime fails): {live_wrong_sorted[:8]}\n\n"
            "qid:",
            text=str(default_qid) if default_qid else "",
        )
        if not ok or not qid_str.strip():
            return
        try:
            qid = int(qid_str.strip())
        except ValueError:
            QMessageBox.warning(self, "Nuclear Review", "qid must be an integer.")
            return
        if qid not in loaded_qids:
            QMessageBox.warning(self, "Nuclear Review", f"Q{qid} is not in the loaded docx.")
            return

        try:
            q = self._q_by_id(qid)
        except KeyError:
            QMessageBox.warning(self, "Nuclear Review", f"Q{qid} not found.")
            return

        sel = sorted(self.selections.get(qid, set()))
        your = ",".join(sel) if sel else "(no answer recorded)"
        item = {
            "qid": q.qid,
            "stem": q.stem,
            "options": dict(q.options),
            "correct_answer": q.effective_answer(self.overrides),
            "your_answer": your,
            "explanation": (self.my_expl.get(q.qid) or q.explanation or ""),
            "pattern_id": q.pattern_id or "",
            "tags": list(q.tags or []),
            "lifetime_wrong_count": wrong_counts.get(qid, 0),
        }

        # Pull relevant concept dossiers (matching the question's text).
        haystack = (q.stem or "") + "\n" + " ".join((q.options or {}).values())
        concepts_hit = _detect_concepts(haystack)
        dossier_dir = self._provider_dossier_dir(provider)
        dossier_chunks: List[str] = []
        for c in concepts_hit:
            slug = _slug_for_concept(c)
            p = dossier_dir / f"{slug}.md"
            if p.exists():
                try:
                    dossier_chunks.append(f"## Dossier: {c}\n\n{p.read_text(encoding='utf-8', errors='replace')}")
                except Exception:
                    pass
        dossier_md = "\n\n---\n\n".join(dossier_chunks)

        # Pull master excerpts (recent prior wrong-answer reports for context)
        master_excerpts: List[Dict] = []
        master_path = self._provider_report_path("coach", provider)
        if master_path.exists():
            try:
                text = master_path.read_text(encoding="utf-8", errors="replace")
                sections = re.split(r"(?=^###\s+Q\d+)", text, flags=re.MULTILINE)
                for sec in sections:
                    m = re.match(r"^###\s+Q(\d+)\b[^\n]*?(?:(?:\u2014|-)\s*(.+?))?\n", sec)
                    if not m:
                        continue
                    excerpt_qid = int(m.group(1))
                    body = sec[m.end():].strip()[:2000]
                    if not body:
                        continue
                    master_excerpts.append({
                        "qid": excerpt_qid,
                        "topic": (m.group(2) or "").strip(),
                        "body": body,
                    })
                # Cap to most recent 30 to keep the synthesizer's context manageable.
                master_excerpts = master_excerpts[-30:]
            except Exception:
                pass

        # Dedup against nuclear_reports.md so we don't re-burn 4 Claude calls
        # on a qid that was already nuclear-reviewed.
        nuclear_path_check = self._provider_report_path("nuclear", provider)
        prior_nuclear_runs = 0
        if nuclear_path_check.exists():
            try:
                _ntext = nuclear_path_check.read_text(encoding="utf-8", errors="replace")
                prior_nuclear_runs = len(re.findall(
                    rf"<!--\s*NUCLEAR_REVIEW\s+qid={qid}\b", _ntext
                ))
            except Exception:
                prior_nuclear_runs = 0

        if prior_nuclear_runs > 0:
            ans = QMessageBox.question(
                self, "\u2622 Nuclear Review",
                f"Q{qid} already exists in {nuclear_path_check.name} ({prior_nuclear_runs} prior run(s)).\n\n"
                f"Re-correr Nuclear (4 {backend_label} calls, ~3-5 min)? Quema tokens.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel,
            )
            if ans != QMessageBox.StandardButton.Yes:
                return

        # Confirm cost.
        ans = QMessageBox.question(
            self, "\u2622 Nuclear Review",
            f"Q{qid} - fan-out 4 {backend_label} calls (3 parallel + 1 synth).\n\n"
            f"Concept dossiers loaded: {len(dossier_chunks)} ({', '.join(concepts_hit) or 'none'}).\n"
            f"History excerpts: {len(master_excerpts)} (last 30 wrong-answer reports).\n\n"
            f"Estimated time: 3-5 min. Proceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        try:
            import quiz_ai_coach as _qac
            _qac._logger.info("nuclear preflight ok (python=%s)", sys.executable)
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, "Nuclear Review - cannot start",
                f"Failed to import quiz_ai_coach.\n\n{type(_e).__name__}: {_e}\n\n{_tb.format_exc()}",
            )
            return

        progress = QProgressDialog(
            f"\u2622 Nuclear Review on Q{qid}...\n"
            f"{backend_label}: spawning 3 parallel agents + synthesizer (3-5 min)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("Nuclear Review")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_nuclear.setEnabled(False)
        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "scope": "nuclear",
            "ai_provider": provider,
            "ai_backend_label": backend_label,
            "pool": [qid],
            "sent_n": 1,
            "sent_qids": {qid},
            "scope_label": f"nuclear \u00b7 Q{qid}",
            "tick": 0,
            "last_msg": "spawning sub-agents...",
            "cancelled": False,
            "is_nuclear": True,
            "nuclear_qid": qid,
            "nuclear_concepts": concepts_hit,
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        self._ai_coach_worker = NuclearReviewWorker(
            item, master_excerpts, dossier_md,
            model=ai_model, provider=provider, timeout_sec=600.0,
        )
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._nuclear_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_thread.start()

    def _nuclear_finished(self, result: dict, error: str) -> None:
        state = dict(self._ai_coach_state or {})
        progress = self._ai_coach_progress
        heartbeat = self._ai_coach_heartbeat
        worker = self._ai_coach_worker
        thread = self._ai_coach_thread
        self._ai_coach_progress = None
        self._ai_coach_heartbeat = None
        self._ai_coach_worker = None
        self._ai_coach_thread = None
        self._ai_coach_state = None
        if heartbeat:
            heartbeat.stop()
            heartbeat.deleteLater()
        if progress:
            progress.close()
            progress.deleteLater()
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        try:
            self.btn_nuclear.setEnabled(True)
        except Exception:
            pass
        if state.get("cancelled"):
            return
        if error:
            QMessageBox.critical(self, "Nuclear Review failed", f"{self._ai_error_for_dialog(error)}\n\nSee ai_coach.log.")
            return
        final = (result or {}).get("final", "").strip()
        if not final:
            QMessageBox.warning(self, "Nuclear Review", "Empty final response. See ai_coach.log.")
            return

        qid = int(state.get("nuclear_qid", 0) or 0)
        concepts = list(state.get("nuclear_concepts", []))
        provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
        nuclear_path = self._provider_report_path("nuclear", provider_for_save)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        section = (
            f"\n\n<!-- NUCLEAR_REVIEW qid={qid} concepts={concepts} ts={ts} -->\n"
            f"# {ts} - Nuclear Review - Q{qid}\n\n"
            f"## Final synthesis\n\n{final}\n\n"
            f"<details>\n<summary>Sub-agent outputs (boundary / patterns / distractors)</summary>\n\n"
            f"### Boundary analysis\n\n{result.get('boundary', '')}\n\n"
            f"### Pattern analysis\n\n{result.get('patterns', '')}\n\n"
            f"### Distractor forensics\n\n{result.get('distractors', '')}\n\n"
            f"</details>\n"
        )
        try:
            with nuclear_path.open("a", encoding="utf-8") as f:
                f.write(section)
            appended = f"Appended to: {nuclear_path.name}"
        except Exception as _ex:
            appended = f"(append failed: {_ex})"
        footer = (
            f"\n\n---\n_Nuclear Review \u00b7 Q{qid} \u00b7 concepts: {', '.join(concepts) or 'none'}_\n\n"
            f"_{appended}_"
        )
        dlg = AICoachDialog(
            self, final + footer, num_questions=1,
            model=f"{state.get('ai_backend_label', 'AI')} \u00b7 nuclear \u00b7 Q{qid}",
            master_path=nuclear_path,
        )
        dlg.setWindowTitle(f"\u2622 Nuclear Review - Q{qid}")
        self._show_coach_dialog_nonmodal(dlg)

    # ------------------------------------------------------------------
    # Diagram - generate a self-contained HTML study artifact for one qid.
    # Reads any prior reports/artifacts for the qid so the model doesn't
    # re-derive from zero. Writes Q<N>_diagram.html and opens in browser.
    # ------------------------------------------------------------------
    def _read_qid_section(self, path: Path, qid: int, marker_re: str) -> str:
        """Return the most recent section in `path` that begins with a marker matching qid."""
        if not path.exists():
            return ""
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return ""
        sections = re.split(rf"(?={marker_re})", text)
        matches: List[str] = []
        for sec in sections:
            m = re.match(marker_re, sec)
            if not m:
                continue
            try:
                sec_qid = int(m.group(1))
            except (ValueError, IndexError):
                continue
            if sec_qid == qid:
                matches.append(sec.strip())
        return matches[-1] if matches else ""

    def _read_ai_coach_section(self, qid: int) -> str:
        """Find the most recent ### Q<qid> section in ai_coach_reports.md."""
        path = self._provider_report_path("coach")
        if not path.exists():
            return ""
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return ""
        sections = re.split(r"(?=^###\s+Q\d+)", text, flags=re.MULTILINE)
        matches: List[str] = []
        for sec in sections:
            m = re.match(r"^###\s+Q(\d+)\b", sec)
            if not m:
                continue
            if int(m.group(1)) == qid:
                matches.append(sec.strip())
        return matches[-1] if matches else ""

    def _read_side_artifacts(self, qid: int) -> str:
        """Concatenate any Q<qid>_*.md / .txt next to the script (excluding diagrams)."""
        try:
            base = Path(__file__).parent
        except Exception:
            base = Path.cwd()
        chunks: List[str] = []
        for ext in ("md", "txt"):
            for p in sorted(base.glob(f"Q{qid}_*.{ext}")):
                # skip the diagram output files themselves
                if p.suffix.lower() == ".html":
                    continue
                try:
                    chunks.append(f"## {p.name}\n\n{p.read_text(encoding='utf-8', errors='replace')}")
                except Exception:
                    pass
        return "\n\n---\n\n".join(chunks)

    def show_rag_kb(self) -> None:
        """Ask the local RAG knowledge base about the current question.

        This is the bridge between the original rich quiz GUI and the new
        enterprise-RAG scaffold in src/network_ai_assistant/rag.
        """
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, "RAG KB", "A coach call is already running. Wait for it to finish.")
            return
        if not self.questions or not self.order:
            QMessageBox.information(self, "RAG KB", "Load a quiz first.")
            return

        try:
            q = self._current_question()
        except Exception:
            QMessageBox.warning(self, "RAG KB", "No current question selected.")
            return

        try:
            from network_ai_assistant.llm.safety import load_env_file
            from network_ai_assistant.rag.domain import SourceDocument
            from network_ai_assistant.rag.ingest import chunk_documents, load_markdown_documents
            from network_ai_assistant.rag.prompting import build_grounded_prompt, deterministic_summary
            from network_ai_assistant.rag.simple_retriever import KeywordRetriever
        except Exception as exc:
            QMessageBox.critical(
                self,
                "RAG KB",
                "Could not import the RAG modules from src/network_ai_assistant.\n\n"
                f"{type(exc).__name__}: {exc}",
            )
            return

        repo_root = _REPO_ROOT if "_REPO_ROOT" in globals() else Path.cwd()
        load_env_file(repo_root / ".env")
        private_kb = repo_root / "data" / "private" / "knowledge_docs.md"
        mock_kb = repo_root / "data" / "mock" / "network_docs.md"
        kb_path = private_kb if private_kb.exists() else (mock_kb if mock_kb.exists() else None)

        sel = sorted(self.selections.get(q.qid, set()))
        your = ",".join(sel) if sel else "(no answer recorded)"
        options_text = "\n".join(f"{letter}) {text}" for letter, text in sorted((q.options or {}).items()))
        current_question_context = (
            f"Q{q.qid}\n\n"
            f"Stem:\n{q.stem}\n\n"
            f"Options:\n{options_text}\n\n"
            f"Your answer: {your}\n"
            f"Correct answer: {q.effective_answer(self.overrides)}\n"
            f"Tags: {', '.join(q.tags or [])}\n"
            f"Pattern: {q.pattern_id or ''}\n"
        )

        def required_focus_terms(text: str) -> List[str]:
            low = (text or "").lower()
            term_groups = [
                ["multicast"],
                ["direct connect", "dx", "transit vif", "private vif"],
                ["route 53 resolver", "resolver endpoint", "dns firewall"],
                ["network firewall", "firewall manager"],
                ["gateway load balancer", "gwlb"],
                ["private link", "privatelink", "endpoint service"],
                ["global accelerator"],
                ["site-to-site vpn", "ipsec", "customer gateway"],
                ["transit gateway", "tgw"],
                ["nat gateway"],
                ["vpc peering"],
                ["load balancer", "nlb", "alb"],
            ]
            out: List[str] = []
            for group in term_groups:
                if any(term in low for term in group):
                    out.extend(group)
            return out

        focus_terms = required_focus_terms(current_question_context)

        def has_focus(result) -> bool:
            if not focus_terms:
                return True
            hay = (result.chunk.text or "").lower()
            # If the current question has a high-signal term such as multicast,
            # require that term in candidate question-bank evidence. This avoids
            # generic VPC/EC2/Security Group matches that feel unrelated.
            if "multicast" in focus_terms:
                return "multicast" in hay
            return any(term in hay for term in focus_terms)

        try:
            quiz_docs: List[SourceDocument] = []
            for qq in self.questions or []:
                if int(qq.qid) == int(q.qid):
                    continue
                qq_options = "\n".join(
                    f"{letter}) {text}" for letter, text in sorted((qq.options or {}).items())
                )
                qq_expl = (self.my_expl.get(qq.qid) or qq.explanation or "").strip()
                qq_text = (
                    f"## Loaded Question Q{qq.qid}\n\n"
                    f"Stem:\n{qq.stem}\n\n"
                    f"Options:\n{qq_options}\n\n"
                    f"Correct answer: {qq.effective_answer(self.overrides)}\n"
                    f"Explanation:\n{qq_expl or '(none)'}\n\n"
                    f"Tags: {', '.join(qq.tags or [])}\n"
                    f"Pattern: {qq.pattern_id or ''}\n"
                )
                quiz_docs.append(
                    SourceDocument(
                        source_id=f"QBANK-Q{qq.qid}",
                        text=qq_text,
                        metadata={
                            "title": f"Loaded Question Q{qq.qid}",
                            "source_type": "loaded_question_bank",
                            "qid": int(qq.qid),
                            "tags": ", ".join(qq.tags or []),
                            "pattern": qq.pattern_id or "",
                        },
                    )
                )

            quiz_chunks = chunk_documents(quiz_docs, max_chars=1400)
            quiz_results = KeywordRetriever(quiz_chunks).search(current_question_context, k=30)
            quiz_results = [r for r in quiz_results if r.score >= 0.08 and has_focus(r)]

            supplemental_results = []
            if kb_path is not None and kb_path.exists():
                docs = load_markdown_documents(kb_path)
                chunks = chunk_documents(docs)
                supplemental_results = KeywordRetriever(chunks).search(current_question_context, k=3)
                supplemental_results = [r for r in supplemental_results if r.score >= 0.06 and has_focus(r)]

            results = quiz_results[:4]
            seen_chunks = {r.chunk.chunk_id for r in results}
            for result in supplemental_results:
                if len(results) >= 6:
                    break
                if result.chunk.chunk_id in seen_chunks:
                    continue
                results.append(result)
                seen_chunks.add(result.chunk.chunk_id)

            kb_label_parts = [f"loaded question bank ({len(quiz_docs)} related candidates)"]
            if kb_path is not None and kb_path.exists():
                kb_label_parts.append(f"supplemental {kb_path.name}")
            kb_label = " + ".join(kb_label_parts)

            rag_question = (
                "Use the retrieved question-bank context to help analyze this current quiz item. "
                "Explain what context applies, what is missing, and which validation/checks "
                "would matter in a real network or AWS environment. If retrieved evidence "
                "contradicts the provided answer key, call that out as a possible wording or "
                "answer-key issue.\n\n"
                f"{current_question_context}"
            )
            grounded_prompt = build_grounded_prompt(rag_question, results)
            summary = deterministic_summary(rag_question, results)
            if not quiz_results:
                summary = (
                    "Context fit warning: no strong related question was retrieved from the loaded bank. "
                    "The answer may need private notes, AWS docs, or a live LLM call with the current question context.\n\n"
                    + summary
                )
        except Exception as exc:
            QMessageBox.critical(self, "RAG KB", f"RAG retrieval failed:\n\n{type(exc).__name__}: {exc}")
            return

        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()
        if not self._ensure_ai_backend_ready():
            return

        ans = QMessageBox.question(
            self,
            "RAG KB",
            f"Q{q.qid} - run one grounded study review with {backend_label}?\n\n"
            "This sends the current question plus retrieved context to the selected AI backend,\n"
            "same style as AI Coach / Deep Review.\n\n"
            f"Knowledge base: {kb_label}\n"
            f"Retrieved chunks: {len(results)}\n\n"
            "Proceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Yes,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        system_prompt = (
            "You are a senior AWS networking study coach and network infrastructure assistant. "
            "Use the current question and retrieved context. Be practical, concise, and educational. "
            "If the retrieved context is weak or contradicts the provided answer key, say that clearly. "
            "Cite retrieved sources as [1], [2], etc."
        )
        user_msg = (
            f"{grounded_prompt}\n\n"
            "Return this structure:\n"
            "1. Quick verdict\n"
            "2. Why the selected answer is wrong or incomplete\n"
            "3. How to reason through each relevant option\n"
            "4. Real-world AWS/network validation checks\n"
            "5. Memory hook\n"
        )

        progress = QProgressDialog(
            f"RAG KB ({backend_label}): grounded review for Q{q.qid}...\n"
            "(one AI call, usually 15-60 sec)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("RAG KB")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_rag_kb.setEnabled(False)
        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "is_rag_kb": True,
            "qid": q.qid,
            "ai_provider": provider,
            "ai_backend_label": backend_label,
            "scope_label": f"Q{q.qid} grounded context",
            "kb_label": kb_label,
            "summary": summary,
            "grounded_prompt": grounded_prompt,
            "master_path": str(Path(self.docx_path) if self.docx_path else (kb_path or repo_root)),
            "tick": 0,
            "last_msg": "retrieval complete; waiting for model...",
            "cancelled": False,
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        self._ai_coach_worker = RagKbWorker(
            system_prompt,
            user_msg,
            model=ai_model,
            provider=provider,
            timeout_sec=300.0,
        )
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._rag_kb_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_thread.start()

    @Slot(str, str)
    def _rag_kb_finished(self, answer_md: str, error: str) -> None:
        state = dict(self._ai_coach_state or {})
        progress = self._ai_coach_progress
        heartbeat = self._ai_coach_heartbeat
        worker = self._ai_coach_worker
        thread = self._ai_coach_thread

        self._ai_coach_progress = None
        self._ai_coach_heartbeat = None
        self._ai_coach_worker = None
        self._ai_coach_thread = None
        self._ai_coach_state = None

        if heartbeat:
            heartbeat.stop()
            heartbeat.deleteLater()
        if progress:
            progress.close()
            progress.deleteLater()
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        try:
            self.btn_rag_kb.setEnabled(True)
        except Exception:
            pass

        if bool(state.get("cancelled", False)):
            return
        if error:
            QMessageBox.critical(self, "RAG KB failed", f"{self._ai_error_for_dialog(error)}\n\nSee ai_coach.log for details.")
            return
        if not answer_md.strip():
            QMessageBox.warning(self, "RAG KB", "Empty response from AI backend. See ai_coach.log.")
            return

        qid = state.get("qid", "?")
        kb_label = state.get("kb_label", "")
        backend_label = state.get("ai_backend_label", "AI")
        summary = state.get("summary", "")
        grounded_prompt = state.get("grounded_prompt", "")
        master_path = Path(str(state.get("master_path") or Path.cwd()))
        report_md = (
            f"# RAG Knowledge Base - Q{qid}\n\n"
            f"**Knowledge base:** `{kb_label}`\n\n"
            f"**Mode:** live GUI backend via `{backend_label}`\n\n"
            "## Retrieved Evidence\n\n"
            f"```text\n{summary}\n```\n\n"
            "## Study Review\n\n"
            f"{answer_md}\n\n"
            "<details>\n<summary>Grounded prompt</summary>\n\n"
            "```text\n"
            f"{grounded_prompt}\n"
            "```\n\n"
            "</details>\n"
        )
        dlg = AICoachDialog(
            self,
            report_md,
            num_questions=1,
            model=f"RAG KB \u00b7 {backend_label}",
            master_path=master_path,
            header_text=f"RAG Knowledge Base - Q{qid}",
        )
        dlg.setWindowTitle(f"RAG KB - Q{qid}")
        self._show_coach_dialog_nonmodal(dlg)

    def show_diagram(self) -> None:
        """Generate a multi-section HTML study diagram for the CURRENT question.

        Pulls prior artifacts (deep / nuclear / ai_coach sections + Q<N>_*.{md,txt})
        as substrate so the model integrates rather than re-derives. Saves to
        Q<N>_diagram.html and opens in the default browser.
        """
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, "Diagram", "A coach call is already running. Wait for it to finish.")
            return
        if not self.questions or not self.order:
            QMessageBox.information(self, "Diagram", "Load a quiz first.")
            return
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()
        artifact_token = self._provider_artifact_token(provider)

        try:
            q = self._current_question()
        except Exception:
            QMessageBox.warning(self, "Diagram", "No current question selected.")
            return

        qid = q.qid
        sel = sorted(self.selections.get(qid, set()))
        your = ",".join(sel) if sel else "(no answer recorded)"
        wrong_counts = {qid_: int(s.get("wrong", 0) or 0) for qid_, s in self.q_stats.items()}

        item = {
            "qid": q.qid,
            "stem": q.stem,
            "options": dict(q.options),
            "correct_answer": q.effective_answer(self.overrides),
            "your_answer": your,
            "explanation": (self.my_expl.get(q.qid) or q.explanation or ""),
            "pattern_id": q.pattern_id or "",
            "tags": list(q.tags or []),
            "lifetime_wrong_count": wrong_counts.get(qid, 0),
        }

        # Pull prior substrate. Each one may be empty - that's fine.
        deep_section = self._read_qid_section(
            self._provider_report_path("deep", provider),
            qid,
            r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)\b[^>]*-->",
        )
        nuclear_section = self._read_qid_section(
            self._provider_report_path("nuclear", provider),
            qid,
            r"<!--\s*NUCLEAR_REVIEW\s+qid=(\d+)\b[^>]*-->",
        )
        ai_coach_section = self._read_ai_coach_section(qid)
        side_artifacts = self._read_side_artifacts(qid)

        prior: Dict[str, str] = {}
        if deep_section:
            prior["deep_review_section"] = deep_section
        if nuclear_section:
            prior["nuclear_section"] = nuclear_section
        if ai_coach_section:
            prior["ai_coach_section"] = ai_coach_section
        if side_artifacts:
            prior["side_artifacts"] = side_artifacts

        # Confirm with substrate inventory so the user knows what we're feeding.
        substrate_lines = []
        substrate_lines.append(f"Deep Review section: {'yes (%d chars)' % len(deep_section) if deep_section else 'none'}")
        substrate_lines.append(f"Nuclear section:     {'yes (%d chars)' % len(nuclear_section) if nuclear_section else 'none'}")
        substrate_lines.append(f"AI Coach section:    {'yes (%d chars)' % len(ai_coach_section) if ai_coach_section else 'none'}")
        substrate_lines.append(f"Side artifacts:      {'yes (%d chars)' % len(side_artifacts) if side_artifacts else 'none'}")
        ans = QMessageBox.question(
            self, "\U0001f4ca Diagram",
            f"Q{qid} - generate HTML study diagram (1 {backend_label} call, ~30-90 sec).\n\n"
            f"Your pick: {your}\n"
            f"Correct:   {item['correct_answer']}\n\n"
            "Substrate to feed the model:\n  - " + "\n  - ".join(substrate_lines) + "\n\n"
            f"Output: Q{qid}_diagram{artifact_token}.html (opens in browser).\n\nProceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Yes,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        try:
            import quiz_ai_coach as _qac  # noqa: F401  (preflight import)
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, "Diagram - cannot start",
                f"Failed to import quiz_ai_coach.\n\n{type(_e).__name__}: {_e}\n\n{_tb.format_exc()}",
            )
            return

        progress = QProgressDialog(
            f"\U0001f4ca Generating diagram for Q{qid}...\nOne {backend_label} call (~30-90 sec)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("Diagram")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_diagram.setEnabled(False)
        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "scope": "diagram",
            "ai_provider": provider,
            "ai_backend_label": backend_label,
            "scope_label": f"diagram \u00b7 Q{qid}",
            "tick": 0,
            "last_msg": "starting...",
            "cancelled": False,
            "is_diagram": True,
            "diagram_qid": qid,
            "diagram_item": item,
            "diagram_substrate": {
                "deep_review": bool(deep_section),
                "nuclear": bool(nuclear_section),
                "ai_coach": bool(ai_coach_section),
                "side_artifacts": bool(side_artifacts),
            },
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        self._ai_coach_worker = DiagramWorker(
            item, prior_artifacts=prior, model=ai_model, provider=provider, timeout_sec=600.0,
        )
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._diagram_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_thread.start()

    def _diagram_finished(self, html_str: str, error: str) -> None:
        state = dict(self._ai_coach_state or {})
        progress = self._ai_coach_progress
        heartbeat = self._ai_coach_heartbeat
        worker = self._ai_coach_worker
        thread = self._ai_coach_thread
        self._ai_coach_progress = None
        self._ai_coach_heartbeat = None
        self._ai_coach_worker = None
        self._ai_coach_thread = None
        self._ai_coach_state = None
        if heartbeat:
            heartbeat.stop()
            heartbeat.deleteLater()
        if progress:
            progress.close()
            progress.deleteLater()
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        try:
            self.btn_diagram.setEnabled(True)
        except Exception:
            pass
        if state.get("cancelled"):
            return
        if error:
            QMessageBox.critical(self, "Diagram failed", f"{self._ai_error_for_dialog(error)}\n\nSee ai_coach.log.")
            return
        if not html_str or not html_str.strip():
            QMessageBox.warning(self, "Diagram", "Empty response. See ai_coach.log.")
            return

        qid = int(state.get("diagram_qid", 0) or 0)
        try:
            base = Path(__file__).parent
        except Exception:
            base = Path.cwd()
        provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
        artifact_token = self._provider_artifact_token(provider_for_save)
        # If a Q<N>_diagram.html already exists, version it with a timestamp
        # so we don't clobber prior diagrams the user may want to keep.
        out_path = base / f"Q{qid}_diagram{artifact_token}.html"
        if out_path.exists():
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            out_path = base / f"Q{qid}_diagram{artifact_token}_{ts}.html"
        try:
            out_path.write_text(html_str, encoding="utf-8")
        except Exception as _ex:
            QMessageBox.critical(self, "Diagram - write failed", f"Could not write {out_path}:\n{_ex}")
            return

        # Also append a consolidated index entry to diagram_reports.md so the
        # diagram lives alongside ai_coach_reports.md / deep_review_reports.md /
        # nuclear_reports.md as part of the per-question audit trail.
        index_path = self._provider_report_path("diagram_index", provider_for_save)
        try:
            index_path = self._provider_report_path("diagram_index", provider_for_save)
            ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            item = state.get("diagram_item") or {}
            stem = (item.get("stem") or "").strip().replace("\n", " ")
            if len(stem) > 240:
                stem = stem[:240] + "..."
            correct = item.get("correct_answer") or item.get("correct") or ""
            your = item.get("your_answer") or ""
            tags = item.get("tags") or ""
            substrate = state.get("diagram_substrate") or {}
            substrate_used = ", ".join(k for k, v in substrate.items() if v) or "none"
            entry_lines = [
                "",
                f"### Q{qid} - {ts_now}",
                "",
                f"- **File:** [{out_path.name}](./{out_path.name})",
                f"- **Correct:** {correct}    **Your answer:** {your}",
            ]
            if tags:
                entry_lines.append(f"- **Tags:** {tags}")
            entry_lines.append(f"- **Substrate used:** {substrate_used}")
            if stem:
                entry_lines.append(f"- **Stem:** {stem}")
            entry_lines.append("")
            if not index_path.exists():
                header = (
                    "# Diagram Reports - ANS-C01\n\n"
                    "Index of HTML study diagrams generated by the \U0001f4ca Diagram button "
                    "(Tier 4). Each entry links to the standalone HTML file in this "
                    "same folder.\n"
                )
                index_path.write_text(header, encoding="utf-8")
            with index_path.open("a", encoding="utf-8") as fh:
                fh.write("\n".join(entry_lines) + "\n")
        except Exception as _ex_idx:
            # Index append is best-effort - don't block the user from seeing
            # the diagram if it fails.
            try:
                _logger_local = logging.getLogger("quiz_ai_coach")
                _logger_local.warning("diagram_reports.md append failed: %s", _ex_idx)
            except Exception:
                pass

        try:
            import webbrowser
            webbrowser.open(out_path.resolve().as_uri())
        except Exception:
            pass

        QMessageBox.information(
            self, "\U0001f4ca Diagram",
            f"Saved: {out_path.name}\nIndex: {index_path.name}\n\nOpened in your default browser.",
        )

    def show_teach_zero(self) -> None:
        """Tier 5 - C\u00e1tedra lecture markdown for the CURRENT question.

        Mirrors show_diagram: pulls the same substrate (deep/nuclear/ai_coach/
        side artifacts) and runs analyze_teach_zero. Result is shown in a popup
        QDialog with rendered markdown, and saved as Q<N>_teach_<ts>.md plus an
        index entry in teach_zero_reports.md.
        """
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, "C\u00e1tedra", "A coach call is already running. Wait for it to finish.")
            return
        if not self.questions or not self.order:
            QMessageBox.information(self, "C\u00e1tedra", "Load a quiz first.")
            return
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()
        artifact_token = self._provider_artifact_token(provider)

        try:
            q = self._current_question()
        except Exception:
            QMessageBox.warning(self, "C\u00e1tedra", "No current question selected.")
            return

        qid = q.qid
        sel = sorted(self.selections.get(qid, set()))
        your = ",".join(sel) if sel else "(no answer recorded)"
        wrong_counts = {qid_: int(s.get("wrong", 0) or 0) for qid_, s in self.q_stats.items()}

        item = {
            "qid": q.qid,
            "stem": q.stem,
            "options": dict(q.options),
            "correct_answer": q.effective_answer(self.overrides),
            "your_answer": your,
            "explanation": (self.my_expl.get(q.qid) or q.explanation or ""),
            "pattern_id": q.pattern_id or "",
            "tags": list(q.tags or []),
            "lifetime_wrong_count": wrong_counts.get(qid, 0),
        }

        deep_section = self._read_qid_section(
            self._provider_report_path("deep", provider),
            qid,
            r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)\b[^>]*-->",
        )
        nuclear_section = self._read_qid_section(
            self._provider_report_path("nuclear", provider),
            qid,
            r"<!--\s*NUCLEAR_REVIEW\s+qid=(\d+)\b[^>]*-->",
        )
        ai_coach_section = self._read_ai_coach_section(qid)
        side_artifacts = self._read_side_artifacts(qid)

        prior: Dict[str, str] = {}
        if deep_section:
            prior["deep_review_section"] = deep_section
        if nuclear_section:
            prior["nuclear_section"] = nuclear_section
        if ai_coach_section:
            prior["ai_coach_section"] = ai_coach_section
        if side_artifacts:
            prior["side_artifacts"] = side_artifacts

        substrate_lines = [
            f"Deep Review section: {'yes (%d chars)' % len(deep_section) if deep_section else 'none'}",
            f"Nuclear section:     {'yes (%d chars)' % len(nuclear_section) if nuclear_section else 'none'}",
            f"AI Coach section:    {'yes (%d chars)' % len(ai_coach_section) if ai_coach_section else 'none'}",
            f"Side artifacts:      {'yes (%d chars)' % len(side_artifacts) if side_artifacts else 'none'}",
        ]
        ans = QMessageBox.question(
            self, "\U0001f393 C\u00e1tedra",
            f"Q{qid} - generate Tier 5 lecture markdown (1 {backend_label} call, ~30-90 sec).\n\n"
            f"Your pick: {your}\n"
            f"Correct:   {item['correct_answer']}\n\n"
            "Substrate to feed the model:\n  - " + "\n  - ".join(substrate_lines) + "\n\n"
            f"Output: Q{qid}_teach{artifact_token}_<ts>.md (popup with rendered markdown).\n\nProceed?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Yes,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        try:
            import quiz_ai_coach as _qac  # noqa: F401
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, "C\u00e1tedra - cannot start",
                f"Failed to import quiz_ai_coach.\n\n{type(_e).__name__}: {_e}\n\n{_tb.format_exc()}",
            )
            return

        progress = QProgressDialog(
            f"\U0001f393 Generating c\u00e1tedra for Q{qid}...\nOne {backend_label} call (~30-90 sec)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("C\u00e1tedra")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_teach_zero.setEnabled(False)
        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "scope": "teach_zero",
            "ai_provider": provider,
            "ai_backend_label": backend_label,
            "scope_label": f"c\u00e1tedra \u00b7 Q{qid}",
            "tick": 0,
            "last_msg": "starting...",
            "cancelled": False,
            "is_teach_zero": True,
            "teach_zero_qid": qid,
            "teach_zero_item": item,
            "teach_zero_substrate": {
                "deep_review": bool(deep_section),
                "nuclear": bool(nuclear_section),
                "ai_coach": bool(ai_coach_section),
                "side_artifacts": bool(side_artifacts),
            },
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        self._ai_coach_worker = TeachZeroWorker(
            item, prior_artifacts=prior, model=ai_model, provider=provider, timeout_sec=600.0,
        )
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._teach_zero_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_thread.start()

    def _teach_zero_finished(self, md_str: str, error: str) -> None:
        state = dict(self._ai_coach_state or {})
        progress = self._ai_coach_progress
        heartbeat = self._ai_coach_heartbeat
        worker = self._ai_coach_worker
        thread = self._ai_coach_thread
        self._ai_coach_progress = None
        self._ai_coach_heartbeat = None
        self._ai_coach_worker = None
        self._ai_coach_thread = None
        self._ai_coach_state = None
        if heartbeat:
            heartbeat.stop()
            heartbeat.deleteLater()
        if progress:
            progress.close()
            progress.deleteLater()
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        try:
            self.btn_teach_zero.setEnabled(True)
        except Exception:
            pass
        if state.get("cancelled"):
            return
        if error:
            QMessageBox.critical(self, "C\u00e1tedra failed", f"{self._ai_error_for_dialog(error)}\n\nSee ai_coach.log.")
            return
        if not md_str or not md_str.strip():
            QMessageBox.warning(self, "C\u00e1tedra", "Empty response. See ai_coach.log.")
            return

        qid = int(state.get("teach_zero_qid", 0) or 0)
        try:
            base = Path(__file__).parent
        except Exception:
            base = Path.cwd()
        provider_for_save = str(state.get("ai_provider") or AI_PROVIDER_CLAUDE)
        artifact_token = self._provider_artifact_token(provider_for_save)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_path = base / f"Q{qid}_teach{artifact_token}_{ts}.md"
        try:
            out_path.write_text(md_str, encoding="utf-8")
        except Exception as _ex:
            QMessageBox.critical(self, "C\u00e1tedra - write failed", f"Could not write {out_path}:\n{_ex}")
            return

        try:
            index_path = self._provider_report_path("teach_index", provider_for_save)
            ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            item = state.get("teach_zero_item") or {}
            stem = (item.get("stem") or "").strip().replace("\n", " ")
            if len(stem) > 240:
                stem = stem[:240] + "..."
            correct = item.get("correct_answer") or ""
            your = item.get("your_answer") or ""
            tags = item.get("tags") or ""
            substrate = state.get("teach_zero_substrate") or {}
            substrate_used = ", ".join(k for k, v in substrate.items() if v) or "none"
            entry_lines = [
                "",
                f"### Q{qid} - {ts_now}",
                "",
                f"- **File:** [{out_path.name}](./{out_path.name})",
                f"- **Correct:** {correct}    **Your answer:** {your}",
            ]
            if tags:
                entry_lines.append(f"- **Tags:** {tags}")
            entry_lines.append(f"- **Substrate used:** {substrate_used}")
            if stem:
                entry_lines.append(f"- **Stem:** {stem}")
            entry_lines.append("")
            if not index_path.exists():
                header = (
                    "# C\u00e1tedra Reports - ANS-C01\n\n"
                    "Index of 'Teach Me From Zero' lecture markdowns generated by the "
                    "\U0001f393 C\u00e1tedra button (Tier 5). Each entry links to the standalone .md "
                    "file in this same folder.\n"
                )
                index_path.write_text(header, encoding="utf-8")
            with index_path.open("a", encoding="utf-8") as fh:
                fh.write("\n".join(entry_lines) + "\n")
        except Exception as _ex_idx:
            try:
                _logger_local = logging.getLogger("quiz_ai_coach")
                _logger_local.warning("teach_zero_reports.md append failed: %s", _ex_idx)
            except Exception:
                pass

        dlg = QDialog(self)
        dlg.setWindowTitle(f"\U0001f393 C\u00e1tedra - Q{qid}")
        dlg.resize(950, 750)
        v = QVBoxLayout(dlg)
        v.setContentsMargins(10, 10, 10, 10)
        view = QTextBrowser(dlg)
        view.setOpenExternalLinks(True)
        try:
            view.setMarkdown(md_str)
        except Exception:
            view.setPlainText(md_str)
        v.addWidget(view, 1)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_open = QPushButton("Open .md file")
        btn_open.clicked.connect(lambda: __import__("webbrowser").open(out_path.resolve().as_uri()))
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_open)
        btn_row.addWidget(btn_close)
        v.addLayout(btn_row)
        dlg.exec()

    def _run_deep_or_prebrief(self, *, mode: str) -> None:
        is_deep = (mode == "deep")
        title = "Deep Review" if is_deep else "Pre-Brief"
        if self._ai_coach_thread and self._ai_coach_thread.isRunning():
            QMessageBox.information(self, title, "A coach call is already running.")
            return
        if not self.questions:
            QMessageBox.information(self, title, "Load a quiz first.")
            return
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()

        loaded_qids = {q.qid for q in self.questions}
        current_all = sorted(self._live_wrong_qids() & loaded_qids)
        wrong_db = self._wrong_db_qids()
        bank_total = len(wrong_db)
        bank_all = sorted(wrong_db & loaded_qids)
        repeat_threshold = max(2, int(self.min_wrong))
        big_threshold = max(repeat_threshold + 1, 3)

        wrong_counts: Dict[int, int] = {
            qid: int(s.get("wrong", 0) or 0)
            for qid, s in self.q_stats.items() if qid in loaded_qids
        }
        repeat_all = sorted([q for q, w in wrong_counts.items() if w >= repeat_threshold])
        big_all = sorted(
            [q for q, w in wrong_counts.items() if w >= big_threshold],
            key=lambda q: (-wrong_counts.get(q, 0), q),
        )

        master_qids = self._load_master_qids()
        deep_qids = self._load_deep_qids()
        no_prior_all = sorted([q for q in current_all if q not in master_qids])

        scope_pools: Dict[str, List[int]] = {
            "big": big_all,
            "repeat": repeat_all,
            "current": current_all,
            "bank": bank_all,
            "no_prior": no_prior_all,
        }

        def _meta(pool: List[int]) -> Dict[str, int]:
            sset = set(pool)
            with_prior = len(sset & master_qids)
            deep_done = len(sset & deep_qids)
            blocking = deep_qids if is_deep else master_qids
            pending = len([q for q in pool if q not in blocking])
            return {
                "total": len(pool),
                "with_prior": with_prior,
                "deep_done": deep_done,
                "pending": pending,
            }

        scopes_meta = {k: _meta(v) for k, v in scope_pools.items()}
        scopes_meta["custom"] = {"total": 0, "with_prior": 0, "deep_done": 0, "pending": 0}

        if all(m["total"] == 0 for m in scopes_meta.values()):
            QMessageBox.information(
                self, title,
                "No wrong answers available yet.\n\nGet some wrong first.",
            )
            return

        default_batch = 3 if is_deep else min(8, self.ai_coach_batch_size)
        picker = DeepReviewPickerDialog(
            self,
            scopes=scopes_meta,
            repeat_threshold=repeat_threshold,
            big_threshold=big_threshold,
            wrong_bank_total=bank_total,
            batch_size_default=default_batch,
            mode_label=title,
        )
        if picker.exec() != QDialog.DialogCode.Accepted or not picker.choice:
            return

        scope = picker.choice
        batch_size = int(picker.batch_size)
        force_redo = bool(picker.force_redo)

        if scope == "custom":
            raw = picker.custom_qids_text.replace(",", " ").split()
            try:
                pool = [int(x) for x in raw if x.strip()]
            except ValueError:
                QMessageBox.warning(self, title, "Custom qids debe ser solo numeros separados por coma o espacio.")
                return
            pool = [q for q in pool if q in loaded_qids]
            if not pool:
                QMessageBox.warning(self, title, "Ninguno de esos qids existe en el docx cargado.")
                return
        else:
            pool = scope_pools.get(scope, [])
        if not pool:
            QMessageBox.information(self, title, f"Scope '{scope}' is empty.")
            return

        # Dedup. Deep blocks against deep_review_reports.md; Pre-Brief blocks
        # against ai_coach_reports.md (a prior AI Coach report already covers it).
        blocking = deep_qids if is_deep else master_qids
        blocking_label = (
            self._provider_report_path("deep", provider).name
            if is_deep else self._provider_report_path("coach", provider).name
        )
        if force_redo:
            pending = list(pool)
        else:
            pending = [q for q in pool if q not in blocking]
            if not pending:
                ans = QMessageBox.question(
                    self, title,
                    f"Los {len(pool)} qids del scope '{scope}' ya estan en {blocking_label} "
                    f"({len(set(pool) & blocking)} hits).\n\nRe-correr anyway? (Quema tokens.)",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                )
                if ans != QMessageBox.StandardButton.Yes:
                    return
                pending = list(pool)

        batch_qids = pending[:batch_size]
        prior_reports = self._extract_prior_reports_for_qids(batch_qids) if is_deep else {}

        items: List[Dict] = []
        for qid in batch_qids:
            try:
                q = self._q_by_id(qid)
            except KeyError:
                continue
            wrong_n = wrong_counts.get(qid, 0)
            sel = sorted(self.selections.get(qid, set()))
            your = ",".join(sel) if sel else "(no answer recorded)"
            items.append({
                "qid": q.qid,
                "stem": q.stem,
                "options": dict(q.options),
                "correct_answer": q.effective_answer(self.overrides),
                "your_answer": your,
                "explanation": (self.my_expl.get(q.qid) or q.explanation or ""),
                "pattern_id": q.pattern_id or "",
                "tags": list(q.tags or []),
                "lifetime_wrong_count": wrong_n,
            })
        if not items:
            QMessageBox.information(self, title, "Nothing to send (all qids skipped).")
            return

        sent_n = len(items)
        sent_qids = {it["qid"] for it in items}
        with_prior = sum(1 for it in items if it["qid"] in prior_reports)
        scope_label = (
            f"{scope} - {'deep' if is_deep else 'prebrief'} batch of {sent_n} "
            f"(qids {batch_qids[0]}-{batch_qids[-1]}"
            + (f" \u00b7 {with_prior}/{sent_n} with prior" if is_deep else "")
            + ")"
        )

        try:
            import quiz_ai_coach as _qac
            _qac._logger.info("%s preflight ok from GUI (python=%s)", title, sys.executable)
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, f"{title} - cannot start",
                f"Failed to import quiz_ai_coach.\n\nPython: {sys.executable}\n\n"
                f"{type(_e).__name__}: {_e}\n\n{_tb.format_exc()}",
            )
            return

        if is_deep:
            # Parallel fan-out per qid: wallclock is the slowest qid, not sum.
            # Per-qid budget is generous (600s); GUI timeout = budget + small overhead.
            dyn_timeout = 660.0
            hint = "(timeout {}s - c\u00e1tedra paralela, ~max(qid) wallclock)".format(int(dyn_timeout))
        else:
            dyn_timeout = min(360.0, 60.0 + 18.0 * sent_n)
            hint = "(timeout {}s - pre-brief barato)".format(int(dyn_timeout))

        progress = QProgressDialog(
            f"{title} ({backend_label}): {scope_label}...\n{hint}",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle(title)
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        if is_deep:
            self.btn_deep_review.setEnabled(False)
        elif hasattr(self, "btn_pre_brief"):
            self.btn_pre_brief.setEnabled(False)

        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "scope": scope,
            "ai_provider": provider,
            "pool": list(pool),
            "sent_n": sent_n,
            "sent_qids": set(sent_qids),
            "scope_label": scope_label,
            "tick": 0,
            "last_msg": "waiting for first chunk...",
            "cancelled": False,
            "ai_backend_label": backend_label,
            "is_deep": is_deep,
            "is_prebrief": (not is_deep),
            "with_prior": with_prior,
            "qid_status": ({int(q): "pending" for q in sent_qids} if is_deep else {}),
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        if is_deep:
            self._ai_coach_worker = DeepReviewWorker(
                items, prior_reports=prior_reports, model=ai_model, provider=provider, timeout_sec=dyn_timeout,
            )
        else:
            self._ai_coach_worker = PreBriefWorker(
                items, model=ai_model, provider=provider, timeout_sec=dyn_timeout,
            )
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._ai_coach_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        if is_deep:
            self._ai_coach_worker.one_done.connect(
                self._deep_review_qid_done,
                Qt.ConnectionType.QueuedConnection,
            )
        self._ai_coach_thread.start()

    def _extract_prior_reports_for_qids(self, qids: List[int]) -> Dict[int, str]:
        """For each qid, find the most recent `### Q<qid>` section in
        ai_coach_reports.md and return its body. Missing qids are simply absent."""
        master_path = self._provider_report_path("coach")
        if not master_path.exists() or not qids:
            return {}
        try:
            text = master_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return {}
        out: Dict[int, str] = {}
        # Split on `### Q<id>` headings; keep the last occurrence per qid.
        pattern = re.compile(r"^###\s+Q(\d+)\b[^\n]*\n", re.MULTILINE)
        matches = list(pattern.finditer(text))
        for i, m in enumerate(matches):
            try:
                qid = int(m.group(1))
            except ValueError:
                continue
            if qid not in qids:
                continue
            start = m.end()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            body = text[start:end].strip()
            # Cap each excerpt to keep the prompt sane.
            if len(body) > 4000:
                body = body[:4000] + "\n\n_(prior report truncated)_"
            out[qid] = body  # later iterations overwrite -> keeps the most recent
        return out

    def show_meta_coach(self) -> None:
        """Read accumulated ai_coach_reports.md and ask Claude to produce a meta-coaching plan."""
        master_path = self._provider_report_path("coach")
        if not master_path.exists():
            QMessageBox.information(
                self, "Meta-Coach",
                "No accumulated reports yet.\n\n"
                "Run AI Coach a few times first - Meta-Coach analyzes the entire history "
                "in ai_coach_reports.md to find recurring brain patterns.",
            )
            return
        try:
            reports_md = master_path.read_text(encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Meta-Coach", f"Could not read {master_path}:\n{e}")
            return
        if not reports_md.strip():
            QMessageBox.information(self, "Meta-Coach", "Reports file is empty.")
            return

        # Cache check: if the source .md hasn't changed since the last meta run,
        # offer to show the cached result instead of re-paying for Claude.
        meta_path = self._provider_report_path("meta")
        hash_path = self._provider_report_path("meta_hash")
        import hashlib as _hashlib
        current_hash = _hashlib.sha256(reports_md.encode("utf-8")).hexdigest()
        cached_hash = ""
        if hash_path.exists():
            try:
                cached_hash = hash_path.read_text(encoding="utf-8").strip()
            except Exception:
                cached_hash = ""
        if cached_hash == current_hash and meta_path.exists():
            box = QMessageBox(self)
            box.setWindowTitle("Meta-Coach - cache hit")
            box.setIcon(QMessageBox.Icon.Information)
            box.setText(
                "El reports.md no ha cambiado desde el ultimo Meta-Coach.\n\n"
                "Tienes el resultado cacheado en disco - no es necesario gastar tokens otra vez."
            )
            btn_show = box.addButton("Mostrar cacheado", QMessageBox.ButtonRole.AcceptRole)
            btn_force = box.addButton("Re-correr (paga tokens)", QMessageBox.ButtonRole.DestructiveRole)
            btn_cancel = box.addButton(QMessageBox.StandardButton.Cancel)
            box.exec()
            clicked = box.clickedButton()
            if clicked is btn_cancel:
                return
            if clicked is btn_show:
                try:
                    cached_md = meta_path.read_text(encoding="utf-8")
                except Exception as _e:
                    QMessageBox.warning(self, "Meta-Coach", f"Could not read cached file:\n{_e}")
                    return
                dlg = AICoachDialog(
                    self, cached_md, num_questions=0,
                    model="cached \u00b7 ai_meta_coach.md (no Claude call)",
                    master_path=meta_path,
                    header_text="Meta-Coach (cacheado - fuente sin cambios)",
                )
                dlg.setWindowTitle("Meta-Coach - cached")
                self._show_coach_dialog_nonmodal(dlg)
                return
            # else: user clicked "Re-correr" -> fall through to live call

        # rough batch count for label
        batch_count = reports_md.count("<!-- AI_COACH_BATCH")
        stats_payload = self._build_meta_stats_payload(reports_md)
        unique_qids = stats_payload["totals"]["unique_qids_in_master"]
        pending_n = stats_payload["totals"]["pending_in_bank"]
        hot_n = len(stats_payload["hot_zone"])
        if not self._ensure_ai_backend_ready():
            return
        provider, ai_model = self._selected_ai_backend()
        backend_label = self._ai_backend_label()

        ans = QMessageBox.question(
            self, "Meta-Coach",
            f"Send {len(reports_md):,} chars across {batch_count} batch(es) "
            f"({unique_qids} unique qids, {hot_n} in hot zone, {pending_n} pending in bank) "
            f"to {backend_label} for a meta-analysis + study plan?\n\n"
            f"This is one big call (no batching) and may take 1-3 minutes.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        # preflight
        try:
            import quiz_ai_coach as _qac
            _qac._logger.info("meta preflight ok from GUI (python=%s)", sys.executable)
        except Exception as _e:
            import traceback as _tb
            QMessageBox.critical(
                self, "Meta-Coach - cannot start",
                f"Failed to import quiz_ai_coach.\n\nPython: {sys.executable}\n\n"
                f"{type(_e).__name__}: {_e}\n\n{_tb.format_exc()}",
            )
            return

        progress = QProgressDialog(
            f"Meta-Coach analyzing {len(reports_md):,} chars / ~{batch_count} batches...\n"
            "(big context call - typically 60-180s)",
            "Cancel", 0, 0, self,
        )
        progress.setWindowTitle("Meta-Coach")
        progress.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()

        self.btn_meta_coach.setEnabled(False)
        self._ai_coach_progress = progress
        self._ai_coach_state = {
            "scope": "meta",
            "ai_provider": provider,
            "ai_backend_label": backend_label,
            "pool": [],
            "sent_n": unique_qids,
            "sent_qids": set(),
            "scope_label": (
                f"meta-coach: {unique_qids} unique qids \u00b7 {batch_count} batches \u00b7 "
                f"{hot_n} hot zone \u00b7 {pending_n} pending"
            ),
            "tick": 0,
            "last_msg": "waiting for first chunk...",
            "cancelled": False,
            "is_meta": True,
            "meta_unique_qids": unique_qids,
            "meta_batch_count": batch_count,
            "meta_pending": pending_n,
            "meta_hot_zone": hot_n,
            "meta_stats_payload": stats_payload,
        }

        self._ai_coach_heartbeat = QTimer(self)
        self._ai_coach_heartbeat.setInterval(1000)
        self._ai_coach_heartbeat.timeout.connect(self._tick_ai_coach_heartbeat)
        self._ai_coach_heartbeat.start()

        self._ai_coach_thread = QThread(self)
        self._ai_coach_worker = MetaCoachWorker(
            reports_md, model=ai_model, provider=provider, timeout_sec=600.0, stats_payload=stats_payload,
        )
        self._ai_coach_worker.moveToThread(self._ai_coach_thread)
        self._ai_coach_thread.started.connect(self._ai_coach_worker.run)
        progress.canceled.connect(self._cancel_ai_coach)
        self._ai_coach_worker.progress_text.connect(
            self._ai_coach_progress_update,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_worker.finished.connect(
            self._ai_coach_finished,
            Qt.ConnectionType.QueuedConnection,
        )
        self._ai_coach_thread.start()

    # ---------- exam submit ----------
    def submit_exam(self) -> None:
        if self.mode != "Exam":
            return
        if not self.questions or not self.order:
            return
        self._maybe_reload_overrides()

        self.submitted = True
        self._tick.stop()

        total = len(self.order)
        correct = 0
        unanswered = 0
        wrong_set: Set[int] = set()

        for i in self.order:
            q = self.questions[i]
            sel = set(self.selections.get(q.qid, set()))
            key = set(q.effective_answer(self.overrides))
            ok = bool(sel) and bool(key) and (sel == key)

            if not sel:
                unanswered += 1
                ok = False

            if ok:
                correct += 1
            else:
                wrong_set.add(q.qid)

        # Keep the 'incorrect last submit' bank aligned with what you'd review (includes unanswered)
        self.last_incorrect = set(wrong_set)

        # Record stats + wrong bank ONCE per session (no inflation on repeated submit)
        if not self._session_results_recorded:
            for i in self.order:
                q = self.questions[i]
                sel = set(self.selections.get(q.qid, set()))
                key = set(q.effective_answer(self.overrides))
                ok = bool(sel) and bool(key) and (sel == key)
                self._bump_stat(q.qid, correct=ok)  # unanswered counts as wrong attempt

            if wrong_set:
                self.wrong_bank |= set(wrong_set)
                self._save_wrong_bank()

            self._session_results_recorded = True

        dlg = ResultsDialog(self, total=total, correct=correct, unanswered=unanswered, seconds_spent=self.seconds_spent)
        dlg.exec()

        self._refresh_nav()
        self._render_current()
        self._schedule_autosave()

    def submit_or_finish(self) -> None:
        """Unified 'finish' button: submits Exam or finishes Practice."""
        if self.mode == "Exam":
            self.submit_exam()
        else:
            self.finish_practice()

    def finish_practice(self) -> None:
        """
        Compute final score for the current Practice session, show Results dialog,
        and (only once) record stats + wrong-bank for items that ended up wrong/unanswered.
        """
        if self.mode != "Practice" or not self.questions or not self.order:
            return
        self._maybe_reload_overrides()

        total = len(self.order)
        correct = 0
        unanswered = 0
        wrong_set: Set[int] = set()

        # Recompute final correctness from the current answer state
        self.practice_attempted = set()
        self.practice_correct = set()

        for i in self.order:
            q = self.questions[i]
            sel = set(self.selections.get(q.qid, set()))
            key = set(q.effective_answer(self.overrides))

            if not sel:
                unanswered += 1
                wrong_set.add(q.qid)
                continue

            self.practice_attempted.add(q.qid)

            if key and sel == key:
                correct += 1
                self.practice_correct.add(q.qid)
            else:
                wrong_set.add(q.qid)

        self.last_incorrect = set(wrong_set)

        if not self._session_results_recorded:
            # In Practice mode, attempts are already recorded on Check (or auto-check).
            # On Finish, only record answers that are currently selected and were not
            # already counted with the same selection.
            for i in self.order:
                q = self.questions[i]
                sel = set(self.selections.get(q.qid, set()))
                if not sel:
                    continue

                cur_scored = frozenset(sel)
                if self._last_scored_sel.get(q.qid) == cur_scored:
                    continue

                self._last_scored_sel[q.qid] = cur_scored
                key = set(q.effective_answer(self.overrides))
                ok = bool(key) and (sel == key)
                self._bump_stat(q.qid, correct=ok)

            if wrong_set:
                self.wrong_bank |= set(wrong_set)
                self._save_wrong_bank()

            self._session_results_recorded = True

        dlg = ResultsDialog(self, total=total, correct=correct, unanswered=unanswered, seconds_spent=self.seconds_spent)
        dlg.exec()

        self._refresh_nav()
        self._render_current()
        self._schedule_autosave()


    # ---------- NEW/progress ----------
    def _update_progress_label(self) -> None:
        """Top-right progress label: 'pos/total' plus NEW remaining (session-only)."""
        total = len(self.order) if self.order else 0
        pos = (self.idx + 1) if total else 0
        label = f"{pos} / {total}"
        new_n = len(getattr(self, "session_new_remaining", set()) or set())
        label += f"   |   \U0001f195 {new_n}"
        self.lbl_progress.setText(label)

    def _sync_new_remaining_for_qid(self, qid: int) -> None:
        """Update session_new_remaining for a qid based on whether it currently has a selection."""
        if qid not in getattr(self, "session_new_initial", set()):
            return
        sel = set(self.selections.get(qid, set()))
        if sel:
            self.session_new_remaining.discard(qid)
        else:
            self.session_new_remaining.add(qid)

    # ---------- nav list ----------
    def _populate_nav(self) -> None:
        self.list_q.clear()
        for i in self.order:
            q = self.questions[i]
            it = QListWidgetItem(f"\u25cb  {q.qid:03d}")
            it.setData(Qt.UserRole, q.qid)
            self.list_q.addItem(it)
        self._refresh_nav()
        self._update_progress_label()

    def _refresh_nav(self) -> None:
        if not self.questions or self.list_q.count() == 0:
            return

        for row in range(self.list_q.count()):
            it = self.list_q.item(row)
            qid = int(it.data(Qt.UserRole))
            q = self._q_by_id(qid)
            sel = self.selections.get(qid, set())
            answered = bool(sel)
            flagged = qid in self.flagged

            correct_state: Optional[bool] = None
            if self.submitted and self.mode == "Exam":
                key = set(q.effective_answer(self.overrides))
                correct_state = (answered and key and set(sel) == key)
            elif self.mode == "Practice" and qid in self.practice_attempted:
                key = set(q.effective_answer(self.overrides))
                correct_state = (qid in self.practice_correct) if key else None

            icon = _status_icon(correct_state, flagged, answered)
            s = self.q_stats.get(qid, {}) or {}
            att = int(s.get('attempts', 0) or 0)
            wrong_n = int(s.get('wrong', 0) or 0)
            cor = int(s.get('correct', 0) or 0)
            suffix = f"  (\u00d7{wrong_n})" if wrong_n else ""
            new_badge = " \U0001f195" if (qid in getattr(self, "session_new_remaining", set())) else ""
            it.setText(f"{icon}  {qid:03d}{suffix}{new_badge}")
            it.setToolTip(f"Q{qid:03d} | wrong: {wrong_n} | correct: {cor} | attempts: {att}")

        self.list_q.setCurrentRow(self.idx)

    def _jump_to_clicked(self, item: QListWidgetItem) -> None:
        qid = int(item.data(Qt.UserRole))
        for row in range(self.list_q.count()):
            if int(self.list_q.item(row).data(Qt.UserRole)) == qid:
                self.idx = row
                self.list_q.setCurrentRow(row)
                self._render_current()
                self._schedule_autosave()
                return

    def _show_coach_dialog_nonmodal(self, dlg: QDialog) -> None:
        """Show an AI Coach / Meta / Deep / DB report dialog non-modally so the user
        can click into questions in the main window without losing the report."""
        if not hasattr(self, "_open_coach_dialogs"):
            self._open_coach_dialogs = []
        self._open_coach_dialogs.append(dlg)

        def _drop(_result, d=dlg):
            try:
                self._open_coach_dialogs.remove(d)
            except ValueError:
                pass
        dlg.finished.connect(_drop)
        dlg.setWindowModality(Qt.WindowModality.NonModal)
        dlg.show()

    def jump_to_qid(self, qid: int) -> bool:
        """Public navigator used by AI Coach reports to deep-link into a question.
        Returns True if the qid was found in the current session order."""
        try:
            qid = int(qid)
        except (TypeError, ValueError):
            return False
        for row in range(self.list_q.count()):
            try:
                if int(self.list_q.item(row).data(Qt.UserRole)) == qid:
                    self.idx = row
                    self.list_q.setCurrentRow(row)
                    self._render_current()
                    self._schedule_autosave()
                    return True
            except Exception:
                continue
        return False

    def prev_q(self) -> None:
        if not self.order:
            return
        self.idx = max(0, self.idx - 1)
        self.list_q.setCurrentRow(self.idx)
        self._render_current()
        self._schedule_autosave()

    def next_q(self) -> None:
        if not self.order:
            return

        # Practice QoL: if enabled, Next will auto-check once (then Next again advances)
        if self.mode == "Practice":
            try:
                if hasattr(self, "chk_auto_next") and self.chk_auto_next.isChecked():
                    if self._auto_check_before_advance():
                        return
            except Exception:
                pass

        self.idx = min(len(self.order) - 1, self.idx + 1)
        self.list_q.setCurrentRow(self.idx)
        self._render_current()
        self._schedule_autosave()

    def _auto_check_before_advance(self) -> bool:
        """Return True if we performed an auto-check and should NOT advance yet."""
        if self.mode != "Practice" or not self.order:
            return False
        self._maybe_reload_overrides()

        try:
            q = self._current_question()
            sel = set(self.selections.get(q.qid, set()))
            key_str = q.effective_answer(self.overrides)
            key = set(key_str) if key_str else set()
        except Exception:
            return False

        if not sel:
            self.txt_feedback.setVisible(True)
            self.txt_feedback.setText("Pick an option first \U0001f642")
            self._refresh_nav()
            return True

        cur = frozenset(sel)
        prev = self._last_checked_sel.get(q.qid)
        if prev == cur:
            return False  # already checked for this exact selection

        ok = bool(key) and (sel == key)

        # Mark as attempted for session score
        self.practice_attempted.add(q.qid)
        if ok:
            self.practice_correct.add(q.qid)
        else:
            self.practice_correct.discard(q.qid)

        # Persist per-question stats (attempt/wrong/correct) once per unique selection
        if self._last_scored_sel.get(q.qid) != cur:
            self._last_scored_sel[q.qid] = cur
            self._bump_stat(q.qid, correct=ok)

        # Remember checked selection (so the next Next advances)
        self._last_checked_sel[q.qid] = cur

        # Auto-flag wrong ones for later review
        if not ok:
            if q.qid not in self.flagged:
                self.flagged.add(q.qid)
                self._save_flagged_bank()

        doc_expl = (q.explanation or "").strip()
        my_expl = (self.my_expl.get(q.qid, "") or "").strip()
        notes = (self.notes.get(q.qid, "") or "").strip()

        expl = my_expl or doc_expl
        if notes:
            expl = (expl + "\n\nNotes:\n" + notes).strip() if expl else ("Notes:\n" + notes)

        def fmt_set(s: Set[str]) -> str:
            return ", ".join(sorted(list(s)))

        msg = []
        msg.append("<h3>" + ("\u2705 Correct" if ok else "\u274c Incorrect") + "</h3>")
        msg.append(f"<b>Your answer:</b> {fmt_set(sel)}<br>")
        msg.append(f"<b>Correct answer:</b> {fmt_set(key)}<br>")
        if expl:
            msg.append("<hr>")
            msg.append("<b>Explanation</b><br>")
            msg.append("<div style='white-space:pre-wrap; font-family: Consolas, Menlo, monospace; font-size: 12px;'>" + html.escape(expl) + "</div>")

        self.txt_feedback.setVisible(True)
        self.txt_feedback.setHtml("".join(msg))
        self._refresh_nav()
        self._render_current(keep_feedback=True)
        self._schedule_autosave()
        return True

    # ---------- question rendering ----------
    def _q_by_id(self, qid: int) -> Question:
        for q in self.questions:
            if q.qid == qid:
                return q
        raise KeyError(qid)

    def _current_question(self) -> Question:
        return self.questions[self.order[self.idx]]

    def _clear_options_ui(self) -> None:
        """Safely clear option rows from the options layout.

        We keep the final stretch item (added in _build_ui) so new rows can be
        inserted just before it.
        """
        # Remove everything except the final stretch/spacer.
        try:
            while self.options_l.count() > 1:
                item = self.options_l.takeAt(0)
                if item is None:
                    break
                w = item.widget()
                if w is not None:
                    try:
                        w.setParent(None)
                    except Exception:
                        pass
                    w.deleteLater()
        except Exception:
            # Never let UI teardown crash the app.
            pass

    def _iter_option_rows(self) -> List[OptionRow]:
        out: List[OptionRow] = []
        for i in range(self.options_l.count() - 1):
            w = self.options_l.itemAt(i).widget()
            if isinstance(w, OptionRow):
                out.append(w)
        return out

    def _render_current(self, keep_feedback: bool = False) -> None:
        if not self.order:
            return

        # Re-entrancy guard: rapid clicks (Start/Restart, nav) can trigger nested renders.
        if getattr(self, '_in_render', False):
            return
        self._in_render = True
        try:
            self._maybe_reload_overrides()

            q = self._current_question()
            total = len(self.order)
            self._update_progress_label()
            self.lbl_qtitle.setText(f"Question {q.qid}  ({self.mode})")

            # flag button state
            self.btn_flag.setText("\u2690" if q.qid in self.flagged else "\u2691")

            # stem + hint (choose count)
            exp_count = q.expected_count()
            hint = []
            s = self.q_stats.get(q.qid, {})
            wrong_n = int(s.get('wrong', 0) or 0)
            if wrong_n:
                hint.append(f"Mistakes: {wrong_n}")
            if exp_count:
                hint.append(f"Choose {exp_count}")
            if self.overrides.get(q.qid):
                hint.append("Overrides \u2705")
            if self.custom_ids_active:
                ids_n = len(self._parse_custom_ids())
                hint.append(f"Custom IDs: {ids_n}")
            if self.mode == "Practice":
                attempted_n = len(self.practice_attempted)
                hint.append(f"Practice score: {len(self.practice_correct)}/{attempted_n} (correct/attempted)")
            self.lbl_hint.setText(" \u00b7 ".join(hint) if hint else " ")

            self.txt_stem.setText(q.stem)

            if self.mode == "Exam" and self.submitted:
                keep_feedback = True

            if not keep_feedback:
                self.txt_feedback.setVisible(False)
                self.txt_feedback.setText("")

            # Determine single vs multi
            key = q.effective_answer(self.overrides)
            is_multi = (exp_count is not None and exp_count > 1) or (len(key) > 1)

            # selections for this q
            sel = set(self.selections.get(q.qid, set()))

            # rebuild options UI
            # Important: drop the previous button group BEFORE deleting old option widgets.
            # Otherwise the group can keep references to buttons that are about to be destroyed.
            if self._option_group is not None:
                try:
                    for b in list(self._option_group.buttons()):
                        try:
                            self._option_group.removeButton(b)
                        except Exception:
                            pass
                    self._option_group.deleteLater()
                except Exception:
                    pass
                self._option_group = None

            self._clear_options_ui()

            # For single-choice questions we enforce exclusivity via a QButtonGroup
            group: Optional[QButtonGroup] = None
            if not is_multi:
                group = QButtonGroup(self.options_host)
                group.setExclusive(True)
                self._option_group = group

            # option order
            letters = self.option_order.get(q.qid) or list(q.options.keys())

            for letter in letters:
                txt = q.options.get(letter, "")
                row = OptionRow(letter, txt, is_multi=is_multi)
                row.set_checked(letter in sel)
                row.toggled.connect(lambda L, checked, qid=q.qid, multi=is_multi: self._on_toggle(qid, L, checked, multi))
                if self.mode == "Exam" and self.submitted:
                    try:
                        row.ctrl.setEnabled(False)
                    except Exception:
                        pass

                self.options_l.insertWidget(self.options_l.count() - 1, row)

                if group and isinstance(row.ctrl, QRadioButton):
                    group.addButton(row.ctrl)


            # Exam reveal: after Submit, show correct/incorrect + correct answer + explanation
            if self.mode == "Exam" and self.submitted:
                try:
                    sel_now = set(self.selections.get(q.qid, set()))
                    key_now = set(q.effective_answer(self.overrides))
                    ok_now = bool(sel_now) and bool(key_now) and (sel_now == key_now)

                    doc_expl = (q.explanation or "").strip()
                    my_expl = (self.my_expl.get(q.qid, "") or "").strip()
                    notes = (self.notes.get(q.qid, "") or "").strip()

                    expl = my_expl or doc_expl
                    if notes:
                        expl = (expl + "\n\nNotes:\n" + notes).strip() if expl else ("Notes:\n" + notes)

                    def fmt_set(s: Set[str]) -> str:
                        return ", ".join(sorted(list(s))) if s else "(blank)"

                    msg = []
                    msg.append("<h3>" + ("\u2705 Correct" if ok_now else "\u274c Incorrect") + "</h3>")
                    msg.append(f"<b>Your answer:</b> {fmt_set(sel_now)}<br>")
                    msg.append(f"<b>Correct answer:</b> {fmt_set(key_now)}<br>")
                    if expl:
                        msg.append("<hr>")
                        msg.append("<b>Explanation</b><br>")
                        msg.append("<div style='white-space:pre-wrap; font-family: Consolas, Menlo, monospace; font-size: 12px;'>" + html.escape(expl) + "</div>")

                    self.txt_feedback.setVisible(True)
                    self.txt_feedback.setHtml("".join(msg))
                except Exception:
                    pass

            self._refresh_nav()
        finally:
            self._in_render = False


    def _on_toggle(self, qid: int, letter: str, checked: bool, is_multi: bool) -> None:
        if is_multi:
            sel = set(self.selections.get(qid, set()))
            if checked:
                sel.add(letter)
            else:
                sel.discard(letter)
            self.selections[qid] = sel
        else:
            # For exclusive (single-choice) groups, ignore the unchecked events
            # that fire when switching between options.
            if not checked:
                return
            self.selections[qid] = {letter}

        self._sync_new_remaining_for_qid(qid)
        self._update_progress_label()
        self._refresh_nav()
        self._schedule_autosave()

    def toggle_flag(self) -> None:
        if not self.order:
            return
        qid = self._current_question().qid
        if qid in self.flagged:
            self.flagged.remove(qid)
        else:
            self.flagged.add(qid)
        self._save_flagged_bank()
        self._refresh_nav()
        self._render_current()
        self._schedule_autosave()

    def clear_selection(self) -> None:
        if not self.order:
            return
        qid = self._current_question().qid
        self.selections[qid] = set()
        self._sync_new_remaining_for_qid(qid)
        self._update_progress_label()
        for row in self._iter_option_rows():
            row.set_checked(False)
        self._refresh_nav()
        self._schedule_autosave()


    # ---------- practice feedback ----------
    def check_practice(self) -> None:
        if self.mode != "Practice" or not self.order:
            return
        self._maybe_reload_overrides()

        try:
            q = self._current_question()
            sel = set(self.selections.get(q.qid, set()))
            key_str = q.effective_answer(self.overrides)
            key = set(key_str) if key_str else set()
        except Exception as e:
            self.txt_feedback.setVisible(True)
            self.txt_feedback.setText(f"Error processing question: {e}")
            return


        if not sel:
            self.txt_feedback.setVisible(True)
            self.txt_feedback.setText("Pick an option first \U0001f642")
            self._refresh_nav()
            return
        # Mark as attempted for the in-session score
        self.practice_attempted.add(q.qid)

        ok = bool(key) and (sel == key)

        # Keep the practice score based on the *current* answer state
        if ok:
            self.practice_correct.add(q.qid)
        else:
            self.practice_correct.discard(q.qid)

        # Persist per-question stats (attempt/wrong/correct) once per unique selection
        cur_scored = frozenset(sel)
        if self._last_scored_sel.get(q.qid) != cur_scored:
            self._last_scored_sel[q.qid] = cur_scored
            self._bump_stat(q.qid, correct=ok)

        # Remember this selection as 'checked' so Next can advance without re-checking
        self._last_checked_sel[q.qid] = frozenset(sel)

        doc_expl = (q.explanation or "").strip()
        my_expl = (self.my_expl.get(q.qid, "") or "").strip()
        notes = (self.notes.get(q.qid, "") or "").strip()

        expl = my_expl or doc_expl
        if notes:
            expl = (expl + "\n\nNotes:\n" + notes).strip() if expl else ("Notes:\n" + notes)

        def fmt_set(s: Set[str]) -> str:
            return ", ".join(sorted(list(s)))

        msg = []
        msg.append("<h3>" + ("\u2705 Correct" if ok else "\u274c Incorrect") + "</h3>")
        msg.append(f"<b>Your answer:</b> {fmt_set(sel)}<br>")
        msg.append(f"<b>Correct answer:</b> {fmt_set(key)}<br>")
        if expl:
            msg.append("<hr>")
            msg.append("<b>Explanation</b><br>")
            msg.append("<div style='white-space:pre-wrap; font-family: Consolas, Menlo, monospace; font-size: 12px;'>" + html.escape(expl) + "</div>")

        self.txt_feedback.setVisible(True)
        self.txt_feedback.setHtml("".join(msg))
        self._refresh_nav()
        self._render_current(keep_feedback=True)  # refresh hint score without clearing feedback
        self._schedule_autosave()


    # ---------- wrong bank ----------
    def clear_wrong_bank(self) -> None:
        if QMessageBox.question(self, "Clear bank", "Clear the saved wrong bank?") == QMessageBox.Yes:
            self.wrong_bank = set()
            self._save_wrong_bank()
            QMessageBox.information(
                self,
                "Done",
                "Saved wrong bank cleared. AI Coach and Deep Review report DB entries remain available."
            )

    # ---------- close ----------
    def closeEvent(self, event):  # type: ignore
        self._save_autosave()
        self._save_paths()
        super().closeEvent(event)


def _apply_dark_theme(app: QApplication) -> None:
    try:
        import qdarkstyle
        app.setStyleSheet(qdarkstyle.load_stylesheet())
        return
    except Exception:
        pass

    app.setStyleSheet(
        """
        QWidget { background: #0e1c2f; color: #e9f0ff; font-family: Segoe UI; }
        QTextBrowser, QListWidget, QComboBox, QSpinBox {
            background: #132944; border: 1px solid #244a73; border-radius: 10px; padding: 6px;
        }
        QPushButton {
            background: #17416a; border: 1px solid #3b6f9e; border-radius: 10px; padding: 8px 10px;
        }
        QPushButton:hover { background: #1f568b; }
        QPushButton:pressed { background: #123353; }
        QToolButton { background: #17416a; border: 1px solid #3b6f9e; border-radius: 10px; padding: 6px 10px; }
        QSplitter::handle { background: #2f5f8f; }
        QSplitter::handle:horizontal { width: 10px; margin: 2px 0; border-left: 1px solid #1a3a5c; border-right: 1px solid #1a3a5c; }
        QSplitter::handle:vertical   { height: 10px; margin: 0 2px; border-top: 1px solid #1a3a5c; border-bottom: 1px solid #1a3a5c; }
        QSplitter::handle:hover { background: #4a86c4; }
        QSplitter::handle:pressed { background: #6aa6e0; }
        """
    )


def main() -> int:
    app = QApplication(sys.argv)
    if USE_DARK_THEME:
        _apply_dark_theme(app)
    w = QuizWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
