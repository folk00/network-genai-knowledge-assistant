"""
Quiz AI Coach - uses claude_agent_sdk to analyze a list of wrong answers
and produce a markdown learning report grouped by concept.

Auth: relies on the local Claude Code install being already authenticated.
No tokens managed here.

Logs: writes detailed timing + state to ai_coach.log next to this script.
"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import traceback
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional
from urllib import error as urllib_error
from urllib import request as urllib_request
from urllib.parse import urlparse

# ---- logging setup -----------------------------------------------------------
_LOG_PATH = Path(__file__).with_name("ai_coach.log")
_logger = logging.getLogger("quiz_ai_coach")
if not _logger.handlers:
    _logger.setLevel(logging.DEBUG)
    _fh = logging.FileHandler(_LOG_PATH, mode="a", encoding="utf-8")
    _fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(threadName)s | %(message)s",
        datefmt="%H:%M:%S",
    ))
    _logger.addHandler(_fh)
_logger.info("=" * 70)
_logger.info("quiz_ai_coach module loaded (pid=%s)", os.getpid())

_CLAUDEGATE_DEFAULT_URL = "http://127.0.0.1:8081"
_CLAUDEGATE_DUMMY_API_KEY = "sk-ant-dummy-key"
OPENAI_DEFAULT_MODEL = "gpt-5.4-mini"

try:
    from claude_agent_sdk import query, ClaudeAgentOptions, AssistantMessage, TextBlock
    _logger.info("claude_agent_sdk imported ok")
except Exception as exc:
    query = ClaudeAgentOptions = AssistantMessage = TextBlock = None  # type: ignore[assignment]
    _CLAUDE_SDK_IMPORT_ERROR = exc
    _logger.exception("Failed to import claude_agent_sdk; Claude provider will be unavailable")
else:
    _CLAUDE_SDK_IMPORT_ERROR = None


SYSTEM_PROMPT_META = (
    "You are a senior cert-prep coach. The user is studying AWS Advanced Networking (ANS-C01) and "
    "has been generating per-question diagnostic reports for every question they got wrong. "
    "You are now given the FULL accumulated history of those reports PLUS a JSON stats sidecar "
    "with reincidence data. USE THE SIDECAR — it tells you which qids matter most.\n\n"
    "Your job: build a META-COACHING plan to fix their brain, not just the answers.\n\n"
    "Output rules:\n"
    "- Output MARKDOWN only. No preamble, no apologies.\n"
    "- Start with `## Brain dump — what's actually broken` — 4-7 ENTRIES (not bullets, full mini-sections). "
    "For each entry, write a `### <pattern name>` heading and then 3-5 paragraphs that:\n"
    "  1. **Name the confusion** in one blunt sentence (e.g. 'You confuse TGW route tables with VPC route tables').\n"
    "  2. **Show the evidence** — cite the specific qids where this fired and quote what the user picked vs the correct answer in 1-2 lines each. Prioritize qids from `stats.hot_zone` first (failed 2+ times with DIFFERENT wrong answers — those mean they're guessing, not memorizing the wrong rule).\n"
    "  3. **Diagnose the misconception** — explain in PROFESSOR TONE why the wrong answer felt right. What mental shortcut is the brain taking? What true-but-incomplete model leads them to the wrong cell? This should be 2-4 sentences of real teaching, not a one-liner.\n"
    "  4. **Install the correct model** — give the actual mental framework the user should be running instead. Use analogies, draw mini ASCII diagrams if helpful, contrast it explicitly with the broken model. This is the rewire part — write like you're tutoring a senior engineer who CAN handle depth, doesn't need basics, but is missing this one piece.\n"
    "  5. **Litmus test** — one sentence: 'Next time you see X in a stem, immediately ask Y'. Make it a reflex they can install.\n"
    "  Each entry should be 200-400 words. Do NOT compress to one bullet. The user explicitly wants elaboration over brevity here.\n"
    "- Then `## Blind spots` — 2-4 bullets on concepts the user keeps missing **even after** "
    "an explanation was already given in a previous batch. Cross-reference `stats.top_reincident`.\n"
    "- Then `## Mental retraining drills` — 4-6 concrete actions (not 'study X', but "
    "'do the AWS docs page on Y, then write a 3-line summary in your own words', "
    "'whiteboard the BGP path selection table from memory', etc.). Each action should target "
    "one of the root causes you named.\n"
    "- Then `## Next session — what to do in the next 30-45 min` — a concrete actionable "
    "list: 6-10 specific qids the user should re-attempt RIGHT NOW (pull from hot_zone first, "
    "then top_reincident), in suggested order, with one-line rationale per qid. This is not a "
    "2-week plan, it's the next 30 minutes. If `stats.pending_in_bank > 0`, also say "
    "'run AI Coach on the N pending qids before this session matters'.\n"
    "- Then `## Study order for the next 2 weeks` — a prioritized list of 5-8 topics with rough "
    "time-boxes (30 min, 2 hrs, etc.) that addresses highest-impact gaps first.\n"
    "- Then `## Flashcards — Anki-style traps` — 5-8 cards in this exact format:\n"
    "  ```\n"
    "  **Card N** — <topic>\n"
    "  - Q: <one-line question that would expose the trap>\n"
    "  - A: <correct answer + 1-line rule>\n"
    "  - Trap: <why the obvious answer is wrong>\n"
    "  ```\n"
    "  Each card targets a specific root cause from the brain dump.\n"
    "- Then `## Exam-day mental checklist` — 5-8 short reminders the user can read 15 min before "
    "the exam to neutralize their specific mistake patterns.\n"
    "- Be blunt. The user is a senior network engineer; do not coddle. Skip basics they obviously know. "
    "If a pattern only appeared once it's probably noise — flag it as 'low priority'.\n"
    "- Reference qids in `[Q123, Q456]` format so the user can navigate back."
)


SYSTEM_PROMPT = (
    "You are a concise, technical AWS Advanced Networking (ANS-C01) tutor. "
    "The user is a senior network engineer studying for the cert and just failed a batch of questions. "
    "Your job: give a per-question diagnosis so they understand exactly why each answer was wrong.\n"
    "\nReport rules:\n"
    "- Output MARKDOWN only. No preamble, no apologies, no 'as an AI'.\n"
    "- For EACH question, emit a `### Q<qid> — <short topic>` heading, then:\n"
    "  - **Question recap:** 2-4 line paraphrase of the stem so the user remembers it without flipping back to the docx. Include the key constraints (e.g. 'TGW with 3 VPCs, on-prem via DX, needs to block VPC-A↔VPC-B'). Then list the options inline as `A) ... · B) ... · C) ... · D) ...` (one short clause each, not the full text).\n"
    "  - **Your answer:** <letter> — one-line why it's tempting but wrong.\n"
    "  - **Correct:** <letter> — one-line why it's right.\n"
    "  - **Key rule:** 1-3 bullets with the AWS rule/mental model the question tests.\n"
    "  - **Trap:** the specific gotcha that made you pick the wrong option.\n"
    "- If a question has a `prior_wrong_answers` field, the user has been analyzed on this qid "
    "before and previously picked one of those letters — they've now flipped to a DIFFERENT wrong "
    "letter. Add an extra bullet `**Flip pattern:**` that calls this out bluntly: name what they "
    "picked then vs. now (e.g. 'Antes B, ahora D — no tienes modelo, estás adivinando'), and tell "
    "them which mental check they need to install so they stop ping-ponging between options. "
    "Treat these qids as HIGH-PRIORITY confusion in the per-batch patterns section.\n"
    "- After all questions, add a `## Patterns in this batch` section: 2-4 bullets clustering "
    "the recurring mistakes (e.g. 'confusing TGW route tables vs VPC route tables').\n"
    "- Be blunt. The user is senior; skip basics. Never invent AWS behavior. If ambiguous, say so."
)


def _build_user_message(items: List[Dict[str, Any]]) -> str:
    payload = {"failed_questions": items}
    return (
        "Here are the questions I just got wrong. Analyze them and produce the report per the rules.\n\n"
        "```json\n"
        + json.dumps(payload, indent=2, ensure_ascii=False)
        + "\n```"
    )


def _normalize_item(raw: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {
        "qid": raw.get("qid"),
        "stem": (raw.get("stem") or "").strip(),
        "options": raw.get("options") or {},
        "correct_answer": (raw.get("correct_answer") or "").strip(),
        "your_answer": (raw.get("your_answer") or "").strip(),
        "explanation": (raw.get("explanation") or "").strip(),
        "pattern_id": (raw.get("pattern_id") or "").strip(),
        "tags": raw.get("tags") or [],
        "lifetime_wrong_count": int(raw.get("lifetime_wrong_count") or 0),
    }
    prior = raw.get("prior_wrong_answers")
    if prior:
        out["prior_wrong_answers"] = list(prior)
    return out


async def _run_query_with_prompt(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
) -> str:
    if query is None or ClaudeAgentOptions is None or AssistantMessage is None or TextBlock is None:
        raise RuntimeError(
            "claude_agent_sdk is required for the Claude provider. "
            f"Install it with `pip install claude-agent-sdk`. Original import error: {_CLAUDE_SDK_IMPORT_ERROR}"
        )
    options = ClaudeAgentOptions(
        system_prompt=system_prompt,
        allowed_tools=[],
        permission_mode="default",
        max_turns=1,
        setting_sources=[],
        mcp_servers={},
    )
    if model:
        options.model = model

    _logger.info("user_msg length=%d chars", len(user_msg))
    chunks: List[str] = []
    msg_count = 0
    block_count = 0
    t0 = time.monotonic()

    async def _consume() -> None:
        nonlocal msg_count, block_count
        async for msg in query(prompt=user_msg, options=options):
            msg_count += 1
            elapsed_msg = time.monotonic() - t0
            _logger.debug("msg #%d type=%s elapsed=%.1fs", msg_count, type(msg).__name__, elapsed_msg)
            try:
                _logger.debug(
                    "msg #%d repr=%s",
                    msg_count,
                    repr(msg)[:1500],
                )
            except Exception:
                _logger.exception("could not repr msg #%d", msg_count)
            for attr in ("subtype", "data", "session_id", "model", "stop_reason", "usage"):
                if hasattr(msg, attr):
                    try:
                        _logger.debug("  msg #%d .%s = %r", msg_count, attr, getattr(msg, attr))
                    except Exception:
                        pass
            if isinstance(msg, AssistantMessage):
                for block in msg.content:
                    if isinstance(block, TextBlock):
                        block_count += 1
                        chunks.append(block.text)
                        elapsed = time.monotonic() - t0
                        _logger.debug(
                            "text block #%d (+%d chars, total=%d, elapsed=%.1fs)",
                            block_count, len(block.text), sum(len(c) for c in chunks), elapsed,
                        )
                        if progress_cb:
                            try:
                                progress_cb(f"received {sum(len(c) for c in chunks)} chars in {elapsed:.0f}s")
                            except Exception:
                                _logger.exception("progress_cb raised")

    _logger.info("starting query (timeout=%.0fs)…", timeout_sec)
    try:
        await asyncio.wait_for(_consume(), timeout=timeout_sec)
    except asyncio.TimeoutError:
        _logger.error(
            "TIMEOUT after %.1fs (msgs=%d, blocks=%d, chars=%d)",
            time.monotonic() - t0, msg_count, block_count, sum(len(c) for c in chunks),
        )
        if chunks:
            return "".join(chunks).strip() + "\n\n---\n_(response truncated by client timeout)_"
        raise RuntimeError(f"Claude did not respond within {timeout_sec:.0f}s. Check ai_coach.log.")

    elapsed = time.monotonic() - t0
    _logger.info(
        "query complete in %.1fs (msgs=%d, blocks=%d, total chars=%d)",
        elapsed, msg_count, block_count, sum(len(c) for c in chunks),
    )
    return "".join(chunks).strip()


def _run_sync(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
) -> str:
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        except Exception:
            _logger.exception("failed to set WindowsProactorEventLoopPolicy")
    loop = asyncio.new_event_loop()
    try:
        asyncio.set_event_loop(loop)
        return loop.run_until_complete(
            _run_query_with_prompt(system_prompt, user_msg, model, progress_cb, timeout_sec)
        )
    except Exception:
        _logger.exception("_run_sync failed")
        raise
    finally:
        try:
            pending = asyncio.all_tasks(loop)
            for t in pending:
                t.cancel()
            if pending:
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        except Exception:
            _logger.exception("error draining loop tasks")
        loop.close()


def _normalize_provider(provider: Optional[str]) -> str:
    p = (provider or "claude").strip().lower()
    if p in {"openai", "gpt", "chatgpt"}:
        return "openai"
    return "claude"


def _redact_api_secrets(text: str) -> str:
    return re.sub(r"\bsk-[A-Za-z0-9_-]{8,}\b", "sk-***", str(text or ""))


def _extract_openai_text(data: Dict[str, Any]) -> str:
    """Extract text from Responses API payloads, with chat-completion fallback."""
    direct = data.get("output_text")
    if isinstance(direct, str) and direct.strip():
        return direct

    chunks: List[str] = []
    for item in data.get("output") or []:
        if not isinstance(item, dict):
            continue
        for content in item.get("content") or []:
            if not isinstance(content, dict):
                continue
            typ = str(content.get("type") or "")
            if typ in {"output_text", "text"}:
                txt = content.get("text")
                if isinstance(txt, str):
                    chunks.append(txt)

    if chunks:
        return "".join(chunks)

    for choice in data.get("choices") or []:
        if not isinstance(choice, dict):
            continue
        msg = choice.get("message") or {}
        if isinstance(msg, dict) and isinstance(msg.get("content"), str):
            chunks.append(msg["content"])
        elif isinstance(choice.get("text"), str):
            chunks.append(choice["text"])
    return "".join(chunks)


def _openai_base_url() -> str:
    base_url = (os.environ.get("OPENAI_BASE_URL") or "https://api.openai.com/v1").strip()
    return base_url.rstrip("/") or "https://api.openai.com/v1"


def _is_local_openai_base(base_url: str) -> bool:
    parsed = urlparse(base_url)
    host = (parsed.hostname or "").lower()
    return host in {"127.0.0.1", "localhost", "::1"}


def _post_openai_json(
    url: str,
    payload: Dict[str, Any],
    api_key: str,
    timeout_sec: float,
) -> Dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib_request.urlopen(req, timeout=timeout_sec) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        _logger.error("OpenAI-compatible endpoint returned non-JSON response: %r", raw[:1200])
        raise RuntimeError("OpenAI-compatible endpoint returned a non-JSON response.") from exc


def _run_openai_sync(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
    *,
    max_output_tokens: int = 16000,
) -> str:
    """Direct OpenAI Responses API call.

    The key is intentionally read only from OPENAI_API_KEY. We do not persist
    API keys in settings or source files.
    """
    base_url = _openai_base_url()
    api_key = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if not api_key and _is_local_openai_base(base_url):
        api_key = "dummy-local-key"
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY is not set. Start this app from a terminal where the "
            "variable exists, set it in your user environment, or point OPENAI_BASE_URL "
            "to a local OpenAI-compatible gateway."
        )

    chosen_model = (model or OPENAI_DEFAULT_MODEL).strip()
    payload = {
        "model": chosen_model,
        "instructions": system_prompt,
        "input": user_msg,
        "max_output_tokens": int(max_output_tokens),
        "store": False,
    }

    t0 = time.monotonic()
    if progress_cb:
        try:
            progress_cb(f"OpenAI-compatible request sent ({chosen_model}); waiting for response...")
        except Exception:
            _logger.exception("progress_cb raised")
    _logger.info(
        "_run_openai_sync: model=%s, base_url=%s, sys_chars=%d, user_chars=%d, timeout=%.0fs, max_output=%d",
        chosen_model, base_url, len(system_prompt), len(user_msg), timeout_sec, int(max_output_tokens),
    )

    try:
        data = _post_openai_json(f"{base_url}/responses", payload, api_key, timeout_sec)
    except urllib_error.HTTPError as exc:
        err_body = exc.read().decode("utf-8", errors="replace")
        redacted = _redact_api_secrets(err_body)
        _logger.error("OpenAI HTTP %s: %s", exc.code, redacted[:2000])
        if _is_local_openai_base(base_url) and exc.code in {400, 404, 405}:
            chat_payload = {
                "model": chosen_model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_msg},
                ],
                "max_tokens": int(max_output_tokens),
            }
            try:
                if progress_cb:
                    progress_cb("Responses endpoint unavailable; retrying chat/completions...")
                data = _post_openai_json(f"{base_url}/chat/completions", chat_payload, api_key, timeout_sec)
            except urllib_error.HTTPError as chat_exc:
                chat_body = chat_exc.read().decode("utf-8", errors="replace")
                redacted_chat = _redact_api_secrets(chat_body)
                _logger.error("OpenAI chat fallback HTTP %s: %s", chat_exc.code, redacted_chat[:2000])
                raise RuntimeError(f"OpenAI-compatible API HTTP {chat_exc.code}: {redacted_chat[:1200]}") from chat_exc
        else:
            raise RuntimeError(f"OpenAI API HTTP {exc.code}: {redacted[:1200]}") from exc
    except urllib_error.URLError as exc:
        _logger.exception("OpenAI connection failed")
        raise RuntimeError(f"OpenAI-compatible API connection failed: {exc}") from exc
    except TimeoutError as exc:
        _logger.exception("OpenAI timed out")
        raise RuntimeError(f"OpenAI-compatible endpoint did not respond within {timeout_sec:.0f}s.") from exc

    text = _extract_openai_text(data).strip()
    elapsed = time.monotonic() - t0
    _logger.info("_run_openai_sync complete in %.1fs (chars=%d)", elapsed, len(text))
    if progress_cb:
        try:
            progress_cb(f"OpenAI returned {len(text)} chars in {elapsed:.0f}s")
        except Exception:
            _logger.exception("progress_cb raised")
    if not text:
        status = data.get("status") or "(unknown)"
        details = data.get("incomplete_details") or data.get("error") or {}
        raise RuntimeError(
            "OpenAI response had no text. "
            f"status={status}; details={_redact_api_secrets(json.dumps(details, ensure_ascii=False))[:800]}"
        )
    return text


def _run_llm_sync(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
    *,
    provider: Optional[str] = "claude",
    max_output_tokens: int = 16000,
) -> str:
    if _normalize_provider(provider) == "openai":
        return _run_openai_sync(
            system_prompt, user_msg, model, progress_cb, timeout_sec,
            max_output_tokens=max_output_tokens,
        )
    return _run_sync(system_prompt, user_msg, model, progress_cb, timeout_sec)


async def _run_llm_async(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
    *,
    provider: Optional[str] = "claude",
    max_output_tokens: int = 16000,
) -> str:
    if _normalize_provider(provider) == "openai":
        return await asyncio.to_thread(
            _run_openai_sync,
            system_prompt,
            user_msg,
            model,
            progress_cb,
            timeout_sec,
            max_output_tokens=max_output_tokens,
        )
    return await _run_query_with_prompt(system_prompt, user_msg, model, progress_cb, timeout_sec)


def _run_visual_or_lecture_sync(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
    *,
    provider: Optional[str] = "claude",
) -> str:
    if _normalize_provider(provider) == "openai":
        return _run_openai_sync(
            system_prompt, user_msg, model, progress_cb, timeout_sec,
            max_output_tokens=16000,
        )
    return _run_diagram_anthropic(system_prompt, user_msg, model, progress_cb, timeout_sec)


def _diagram_gate_base_url() -> str:
    """Return the Claude gate URL used by the Diagram tier."""
    env_url = (os.environ.get("ANTHROPIC_BASE_URL") or "").strip()
    if env_url:
        return env_url.rstrip("/")

    cfg_path = Path.home() / ".config" / "claudegate" / "server.json"
    try:
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        cfg_url = str(cfg.get("url") or "").strip()
        if cfg_url:
            return cfg_url.rstrip("/")
    except Exception:
        _logger.debug("claudegate server.json unavailable; using default", exc_info=True)

    return _CLAUDEGATE_DEFAULT_URL


def _can_connect(host: str, port: int, timeout_sec: float = 0.5) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout_sec):
            return True
    except OSError:
        return False


def _ensure_diagram_gate_running(
    base_url: str,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> None:
    """Start claudegate on demand when the Diagram tier targets localhost."""
    parsed = urlparse(base_url)
    host = parsed.hostname
    if host not in {"127.0.0.1", "localhost"}:
        return

    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    if _can_connect(host, port):
        return

    claudegate = shutil.which("claudegate")
    if not claudegate:
        raise RuntimeError(
            f"Local Claude gate is not listening at {base_url}, and claudegate is not in PATH."
        )

    if progress_cb:
        try:
            progress_cb("starting local Claude gate...")
        except Exception:
            _logger.exception("progress_cb raised")

    _logger.info("starting claudegate for diagram tier: %s", claudegate)
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.run(
        [claudegate, "start"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        timeout=30,
        check=False,
        creationflags=creationflags,
    )
    _logger.info("claudegate start exited rc=%s output=%r", proc.returncode, proc.stdout[-1200:])

    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        if _can_connect(host, port):
            return
        time.sleep(0.5)

    raise RuntimeError(
        f"Local Claude gate did not start at {base_url}. Run `claudegate start` and retry Diagram."
    )


def _run_diagram_anthropic(
    system_prompt: str,
    user_msg: str,
    model: Optional[str],
    progress_cb: Optional[Callable[[str], None]],
    timeout_sec: float,
) -> str:
    """Direct Anthropic SDK call for the Diagram tier.

    Bypasses claude_agent_sdk (which wedges with no AssistantMessage).
    Honors ANTHROPIC_BASE_URL so the local Claude gate continues to handle
    auth/billing exactly as before. Uses streaming + prompt caching on the
    static system prompt.
    """
    import anthropic

    chosen_model = model or "claude-opus-4-7"
    is_opus = chosen_model.startswith("claude-opus")
    base_url = _diagram_gate_base_url()
    _ensure_diagram_gate_running(base_url, progress_cb)
    _logger.info(
        "_run_diagram_anthropic: model=%s, base_url=%s, sys_chars=%d, user_chars=%d, timeout=%.0fs, opus=%s",
        chosen_model,
        base_url,
        len(system_prompt),
        len(user_msg),
        timeout_sec,
        is_opus,
    )

    client = anthropic.Anthropic(
        timeout=timeout_sec,
        base_url=base_url,
        api_key=os.environ.get("ANTHROPIC_API_KEY") or _CLAUDEGATE_DUMMY_API_KEY,
    )
    chunks: List[str] = []
    t0 = time.monotonic()
    last_progress = t0

    stream_kwargs: Dict[str, Any] = dict(
        model=chosen_model,
        max_tokens=24000,
        system=[
            {
                "type": "text",
                "text": system_prompt,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[{"role": "user", "content": user_msg}],
    )
    if is_opus:
        stream_kwargs["thinking"] = {"type": "adaptive"}
        stream_kwargs["output_config"] = {"effort": "medium"}

    try:
        with client.messages.stream(**stream_kwargs) as stream:
            for text_piece in stream.text_stream:
                chunks.append(text_piece)
                now = time.monotonic()
                if progress_cb and (now - last_progress) >= 1.5:
                    last_progress = now
                    total = sum(len(c) for c in chunks)
                    try:
                        progress_cb(f"received {total} chars in {now - t0:.0f}s")
                    except Exception:
                        _logger.exception("progress_cb raised")
            final = stream.get_final_message()
    except Exception as exc:
        elapsed = time.monotonic() - t0
        _logger.exception(
            "_run_diagram_anthropic failed after %.1fs (got %d chars)",
            elapsed, sum(len(c) for c in chunks),
        )
        if chunks:
            return "".join(chunks).strip() + f"\n\n---\n_(stream error after {elapsed:.0f}s: {type(exc).__name__})_"
        raise

    elapsed = time.monotonic() - t0
    usage = getattr(final, "usage", None)
    _logger.info(
        "_run_diagram_anthropic complete in %.1fs (chars=%d, usage=%r, stop=%r)",
        elapsed, sum(len(c) for c in chunks), usage, getattr(final, "stop_reason", None),
    )
    return "".join(chunks).strip()


SYSTEM_PROMPT_DEEP = (
    "You are the ULTIMATE AWS Advanced Networking (ANS-C01) tutor. Tone: senior staff engineer "
    "lecturing another senior. The user is a senior network engineer who keeps screwing the SAME "
    "concept up. Your job for each question is to build the truth FROM ZERO and force a permanent "
    "rewire — not a summary, a full lecture. Length is fine. Density beats brevity.\n\n"
    "You may receive a `prior_report` field per item (a previous brief AI Coach diagnosis or "
    "Pre-Brief context dossier). When present, USE it as context but do not just repeat — you are "
    "the deeper layer. When absent, build from scratch with no excuse.\n\n"
    "Output rules:\n"
    "- Output MARKDOWN only. No preamble, no apologies, no 'sure here is'.\n"
    "- For EACH question, emit `### Q<qid> — <short topic>` and then THESE 7 SECTIONS, in order, "
    "with the exact bold labels shown:\n\n"
    "  **0. Question recap**\n"
    "    - 3-6 line paraphrase of the stem keeping every constraint (services, topology, scale, "
    "what's blocked/required, latency/throughput numbers).\n"
    "    - Then the options, one per line: `- A) <full text>` ... mark the user's pick with "
    "`<- your answer` and the correct one with `<- correct`.\n\n"
    "  **1. Why you screwed it up**\n"
    "    - ONE blunt sentence naming the broken mental model. No hedging.\n"
    "    - Then 1 short paragraph: which true-but-incomplete rule the brain anchored on, and why "
    "it doesn't apply here.\n\n"
    "  **2. Concept map from zero**\n"
    "    - Sub-bullet **AWS services in play:** for each relevant service, one line with what "
    "it solves AND one line with what it does NOT solve (the boundary that catches people).\n"
    "    - Sub-bullet **Networking primitives at play:** the underlying mechanics (BGP attrs, MTU, "
    "ENI placement, route table precedence, AZ failure domain, eventual consistency, etc.) — only "
    "the ones that actually matter for THIS question.\n"
    "    - Sub-bullet **Hard rules people confuse:** 2-4 bullets of crisp 'X NOT Y' "
    "rules. Cite limits, defaults, precedences, who-wins.\n\n"
    "  **3. Step-by-step reconstruction**\n"
    "    - Walk constraint-by-constraint. Format each step as `Constraint: <quote/paraphrase> → "
    "eliminates: <which options and why>`. Build to the correct answer by ELIMINATION, not by "
    "asserting it. The user must be able to re-derive the answer alone next time.\n\n"
    "  **4. Distractor dissection**\n"
    "    - One bullet per option (A, B, C, D, E if present). Format: `- X) <one-line gist>: <why "
    "it's a trap | why it's correct>`. For traps, name the SPECIFIC misconception each one "
    "exploits (anchoring, recency from another question, name-similarity to a real service, etc.).\n\n"
    "  **5. Mental model to install**\n"
    "    - The reusable framework — not specific to this question. 3-5 bullets or a short "
    "numbered procedure. The user should be able to apply it to a NEW question on the same "
    "concept and get it right.\n\n"
    "  **6. Litmus test + radar**\n"
    "    - One reflex: `When you see <trigger> in the stem → immediately <action>`.\n"
    "    - Then 2-3 bullets: **Related questions you'd probably miss with the same error** "
    "(concrete scenario sketches, not real qids — describe the trap shape).\n\n"
    "- After ALL questions, add `## Common thread` — 2-4 bullets connecting the misconceptions if "
    "a pattern recurs. If they're unrelated, say so in one line.\n\n"
    "Hard rules:\n"
    "- Never invent AWS behavior. If genuinely ambiguous or version-dependent, say so explicitly.\n"
    "- Skip 101 basics. Assume CCIE-level networking knowledge; the gap is AWS-specific nuance.\n"
    "- Use small ASCII diagrams when topology matters. Use tables when comparing services.\n"
    "- Be blunt about the user's error. This is a rewire, not validation.\n"
    "- Respond ENTIRELY in English. No Spanish words, headings, or markers anywhere."
)


SYSTEM_PROMPT_PREBRIEF = (
    "You are a cheap, fast pre-analyzer for AWS ANS-C01 questions the user got wrong. Your job is "
    "NOT to teach — that's Deep Review's job later. Your job is to produce a tight CONTEXT DOSSIER "
    "per qid so the deeper analysis has substrate to expand from.\n\n"
    "Output MARKDOWN only. For EACH question, emit `### Q<qid> — <short topic>` then exactly:\n\n"
    "  **Question recap:** 2-4 line paraphrase + options inline as "
    "`A) ... · B) ... · C) ... · D) ...` (mark user's pick with `(you)` and correct with `(ok)`).\n\n"
    "  **Concepts at play:** comma-separated list of 3-6 specific AWS/networking concepts this "
    "question tests (e.g. 'TGW route propagation, BGP AS-path prepending, VPC peering transitivity').\n\n"
    "  **Typical confusion vectors:** 2-3 bullets naming the SHAPES of mistake people make on "
    "this concept (one line each, no full explanations — those come in Deep Review).\n\n"
    "  **What Deep Review must lock down:** 1-3 bullets listing the specific rules/limits/"
    "defaults the deeper lecture must nail down for this qid.\n\n"
    "Hard rules: be terse. No teaching, no paragraphs, no examples. This is a scaffolding pass."
)


def analyze_wrong_answers(
    items: List[Dict[str, Any]],
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    max_items: int = 25,
    timeout_sec: float = 180.0,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> str:
    """Synchronous wrapper. Safe to call from a non-main thread (e.g. QThread)."""
    _logger.info(
        "analyze_wrong_answers called: items=%d, max_items=%d, provider=%s, model=%s",
        len(items), max_items, _normalize_provider(provider), model or "<default>",
    )
    if not items:
        return "_No wrong answers to analyze yet. Get some questions wrong first._"
    normalized = [_normalize_item(x) for x in items[:max_items]]
    user_msg = _build_user_message(normalized)
    return _run_llm_sync(
        SYSTEM_PROMPT, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )


def analyze_meta_coach(
    reports_md: str,
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    max_chars: int = 180_000,
    timeout_sec: float = 600.0,
    progress_cb: Optional[Callable[[str], None]] = None,
    stats_payload: Optional[Dict[str, Any]] = None,
) -> str:
    """Take the accumulated ai_coach_reports.md and produce a meta-coaching plan.

    If the file is huge, the OLDEST content is trimmed (we keep the most recent reports
    since those reflect current weaknesses). `stats_payload` is a JSON-serializable dict
    with reincidence info (hot_zone, top_reincident, totals) injected as a sidecar.
    """
    _logger.info(
        "analyze_meta_coach called: reports_chars=%d, max_chars=%d, provider=%s, model=%s, stats_keys=%s",
        len(reports_md), max_chars, _normalize_provider(provider), model or "<default>",
        list((stats_payload or {}).keys()),
    )
    text = (reports_md or "").strip()
    if not text:
        return "_No accumulated reports yet. Run AI Coach a few times first._"
    truncated_note = ""
    if len(text) > max_chars:
        # keep the tail (most recent batches)
        text = text[-max_chars:]
        truncated_note = (
            f"\n\n_(NOTE: reports file was {len(reports_md)} chars; trimmed oldest content "
            f"to fit {max_chars} chars — analysis is on the most recent batches only.)_\n\n"
        )
        # try to start cleanly at a batch boundary
        marker = text.find("<!-- AI_COACH_BATCH")
        if marker > 0:
            text = text[marker:]

    stats_block = ""
    if stats_payload:
        stats_block = (
            "Stats sidecar (USE THIS to prioritize):\n"
            "```json\n"
            + json.dumps(stats_payload, indent=2, ensure_ascii=False)
            + "\n```\n\n"
        )

    user_msg = (
        "Below is the full accumulated AI Coach report history plus a stats sidecar. "
        "Analyze them per the rules.\n\n"
        f"{stats_block}"
        f"{truncated_note}"
        "Reports history:\n"
        "```markdown\n"
        f"{text}\n"
        "```"
    )
    return _run_llm_sync(
        SYSTEM_PROMPT_META, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )


def analyze_deep_review(
    items: List[Dict[str, Any]],
    *,
    prior_reports: Optional[Dict[int, str]] = None,
    model: Optional[str] = None,
    provider: str = "claude",
    max_items: int = 15,
    timeout_sec: float = 600.0,
    progress_cb: Optional[Callable[[str], None]] = None,
    parallel: bool = True,
    on_qid_done: Optional[Callable[[int, str], None]] = None,
) -> str:
    """Per-question deep mini-class. Optionally expands prior AI Coach reports.

    `prior_reports` maps qid -> the previous brief markdown excerpt for that qid
    (extracted from ai_coach_reports.md by the caller). Missing qids get a fresh
    deep analysis.

    When `parallel=True` (default), each qid is dispatched to its own agent via
    asyncio.gather — wallclock ≈ slowest qid, not sum. Same token cost.
    """
    _logger.info(
        "analyze_deep_review called: items=%d, with_prior=%d, provider=%s, model=%s, parallel=%s",
        len(items), len(prior_reports or {}), _normalize_provider(provider), model or "<default>", parallel,
    )
    if not items:
        return "_No questions to deep-review._"
    normalized = [_normalize_item(x) for x in items[:max_items]]
    prior_reports = prior_reports or {}
    for it in normalized:
        prior = prior_reports.get(it.get("qid"))
        if prior:
            it["prior_report"] = prior.strip()

    def _emit(s: str) -> None:
        if progress_cb:
            try:
                progress_cb(s)
            except Exception:
                pass

    if not parallel:
        payload = {"failed_questions": normalized}
        user_msg = (
            "Here are the wrong questions to deep-review. For any item with a `prior_report` field, "
            "EXPAND on it (do not just repeat). For items without one, write the full mini-class from "
            "scratch. Follow the per-question template strictly.\n\n"
            "```json\n"
            + json.dumps(payload, indent=2, ensure_ascii=False)
            + "\n```"
        )
        return _run_llm_sync(
            SYSTEM_PROMPT_DEEP, user_msg, model, progress_cb, timeout_sec,
            provider=provider,
        )

    async def _per_qid(it: Dict[str, Any]) -> str:
        qid = it.get("qid")
        payload = {"failed_questions": [it]}
        user_msg = (
            "Here is ONE wrong question to deep-review. If a `prior_report` field is present, "
            "EXPAND on it (do not just repeat). Otherwise write the full mini-class from scratch. "
            "Follow the per-question template strictly.\n\n"
            "```json\n"
            + json.dumps(payload, indent=2, ensure_ascii=False)
            + "\n```"
        )
        return await _run_llm_async(
            SYSTEM_PROMPT_DEEP, user_msg, model,
            (lambda s, q=qid: _emit(f"Q{q}: {s}")) if progress_cb else None,
            timeout_sec,
            provider=provider,
        )

    async def _per_qid_safe(it: Dict[str, Any]) -> str:
        """Wrap _per_qid: catches exceptions, fires on_qid_done immediately on completion
        so the GUI can persist results incrementally and never lose work mid-batch."""
        qid = it.get("qid")
        try:
            md = await _per_qid(it)
        except Exception as e:
            _logger.exception("deep: Q%s failed", qid)
            md = f"### Q{qid}\n\n_Error: {type(e).__name__}: {e}_\n"
        if on_qid_done is not None:
            try:
                on_qid_done(int(qid) if qid is not None else 0, md)
            except Exception:
                _logger.exception("on_qid_done raised for Q%s", qid)
        return md

    async def _fanout() -> str:
        _emit(f"deep: spawning {len(normalized)} parallel sub-agents…")
        tasks = [_per_qid_safe(it) for it in normalized]
        results = await asyncio.gather(*tasks)  # exceptions already caught inside _per_qid_safe
        _emit("deep: all sub-agents done.")
        return "\n\n---\n\n".join(str(r).strip() for r in results)

    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        except Exception:
            _logger.exception("failed to set WindowsProactorEventLoopPolicy")
    loop = asyncio.new_event_loop()
    try:
        asyncio.set_event_loop(loop)
        return loop.run_until_complete(_fanout())
    finally:
        try:
            pending = asyncio.all_tasks(loop)
            for t in pending:
                t.cancel()
            if pending:
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        except Exception:
            _logger.exception("error draining loop tasks")
        loop.close()


def analyze_pre_brief(
    items: List[Dict[str, Any]],
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    max_items: int = 25,
    timeout_sec: float = 240.0,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> str:
    """Cheap pre-analysis dossier per qid. Output goes to ai_coach_reports.md so
    Deep Review picks it up later as `prior_report`."""
    _logger.info(
        "analyze_pre_brief called: items=%d, provider=%s, model=%s",
        len(items), _normalize_provider(provider), model or "<default>",
    )
    if not items:
        return "_Nothing to pre-brief._"
    normalized = [_normalize_item(x) for x in items[:max_items]]
    payload = {"failed_questions": normalized}
    user_msg = (
        "Produce a context dossier per qid. NO teaching, NO paragraphs. Follow the template "
        "strictly so Deep Review can use this as scaffolding later.\n\n"
        "```json\n"
        + json.dumps(payload, indent=2, ensure_ascii=False)
        + "\n```"
    )
    return _run_llm_sync(
        SYSTEM_PROMPT_PREBRIEF, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )


SYSTEM_PROMPT_DOSSIER = (
    "You are a senior AWS Advanced Networking instructor consolidating SCATTERED PRIOR TEACHING "
    "into one canonical dossier per concept. The user has been generating per-question deep "
    "reviews that touch this concept from multiple angles. Your job is NOT to invent new content "
    "— you reorganize, deduplicate, and harden what's already been written into a single "
    "reference document the user can study standalone.\n\n"
    "You receive: a concept name + a collection of excerpts (per-question teachings) where this "
    "concept appeared. Each excerpt was written for a specific question; you extract the GENERAL "
    "teaching about the concept and discard the question-specific reasoning.\n\n"
    "Output MARKDOWN only. Structure:\n\n"
    "  # <Concept name>\n\n"
    "  **One-line essence:** the single sentence that captures what this concept actually is "
    "(not what it sounds like).\n\n"
    "  ## What it solves\n"
    "  2-4 bullets, concrete. Each bullet: the use case + why this concept beats alternatives.\n\n"
    "  ## What it does NOT solve (the boundary)\n"
    "  3-6 bullets — the misconceptions, the limits, what people wrongly assume it does. Cite "
    "specific AWS limits/defaults/precedences when relevant.\n\n"
    "  ## How it actually works\n"
    "  The mechanics, in plain technical English. Use small ASCII diagrams when topology helps. "
    "Use tables when comparing related services. Cover the underlying primitives that matter "
    "(routing, BGP attrs, MTU, AZ failure domains, eventual consistency, IAM scope, etc.). "
    "This is the meat of the dossier — be thorough, no fluff.\n\n"
    "  ## Common confusions (specific traps to never fall for again)\n"
    "  3-6 named traps in this format:\n"
    "  - **Trap name:** one-sentence description of the mistake. **Reality:** the correct rule.\n\n"
    "  ## Decision framework\n"
    "  3-5 bullet 'when to use this vs alternative X' rules. Concrete triggers from a question stem "
    "to this concept's answer. The user should be able to apply this to a NEW question and pick "
    "the right service immediately.\n\n"
    "  ## Quick reference card\n"
    "  3-6 one-liner facts that are easy to memorize: limits, default values, what wins on conflict. "
    "Format: `- <fact>`. This is the cheat sheet.\n\n"
    "Hard rules:\n"
    "- Reorganize and consolidate — do NOT invent AWS behavior the excerpts don't support.\n"
    "- Deduplicate ruthlessly: if 5 excerpts said the same thing, write it once and well.\n"
    "- Skip the question-specific reasoning from the excerpts (correct answer was X, you picked Y) "
    "— that belongs in deep_review_reports.md, not here. Keep only the general teaching.\n"
    "- Senior engineer audience. No basics. Density beats brevity, but brevity beats bloat.\n"
    "- This dossier will be reused across MANY future questions on this concept — write it as "
    "permanent reference material."
)


def analyze_concept_dossier(
    concept_name: str,
    source_excerpts: List[Dict[str, Any]],
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    timeout_sec: float = 360.0,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> str:
    """Consolidate scattered teaching about a single concept into one canonical dossier."""
    _logger.info(
        "analyze_concept_dossier: concept=%r, excerpts=%d, provider=%s, model=%s",
        concept_name, len(source_excerpts), _normalize_provider(provider), model or "<default>",
    )
    if not source_excerpts:
        return f"_No source material for {concept_name}._"
    payload = {"concept": concept_name, "source_excerpts": source_excerpts}
    user_msg = (
        f"Concept to consolidate: **{concept_name}**\n\n"
        f"Below are {len(source_excerpts)} excerpts from prior per-question teachings where this "
        "concept appeared. Reorganize them into ONE canonical dossier per the template. "
        "Discard question-specific reasoning; keep the general teaching about the concept.\n\n"
        "```json\n"
        + json.dumps(payload, indent=2, ensure_ascii=False)
        + "\n```"
    )
    return _run_llm_sync(
        SYSTEM_PROMPT_DOSSIER, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )


# =============================================================================
# Nuclear Review — multi-agent fan-out for one critical question.
# Spawns 3 specialist sub-agents in parallel, then a synthesizer combines them
# into the absolute-final-word lecture. Use only on questions you keep missing.
# =============================================================================

SYSTEM_PROMPT_NUCLEAR_BOUNDARY = (
    "You are an AWS Service-Boundary Expert. You receive ONE AWS Advanced Networking question and "
    "an optional concept dossier as reference material. Your sole job: for EACH AWS service "
    "mentioned (or implied by) the stem and options, produce a precise boundary specification.\n\n"
    "Output MARKDOWN only. Structure:\n\n"
    "  ## Service boundaries — Q<qid>\n\n"
    "  For each service, emit:\n\n"
    "  ### <Service name>\n"
    "  - **Solves:** 2-3 bullets — what this service is actually for in this context.\n"
    "  - **Does NOT solve:** 2-4 bullets — adjacent capabilities people wrongly assume.\n"
    "  - **Hard limits/defaults relevant here:** 1-3 specific numbers or rules from AWS docs.\n"
    "  - **Interaction with other services in this stem:** 1-3 bullets if multiple services interact.\n\n"
    "Hard rules:\n"
    "- Only services from the stem/options. Don't invent extra services.\n"
    "- Senior engineer audience. Skip basics. Cite specific AWS limits/defaults.\n"
    "- This output feeds a synthesizer — be precise, not pretty."
)

SYSTEM_PROMPT_NUCLEAR_PATTERNS = (
    "You are a Mistake-Pattern Analyst. You receive ONE AWS Advanced Networking question the user "
    "got wrong, plus their HISTORY of prior wrong-answer reports. Your sole job: identify what "
    "this specific failure says about persistent patterns in their thinking.\n\n"
    "Output MARKDOWN only. Structure:\n\n"
    "  ## Pattern analysis — Q<qid>\n\n"
    "  ### Similar prior failures\n"
    "  Cite 2-5 prior qids from the history that involved the same root misconception. For each: "
    "`Q<id>` (1-line topic) — `[the misconception they share]`. If no clear prior matches, say so.\n\n"
    "  ### What this user's brain actually does wrong (across history)\n"
    "  3-5 bullets identifying THE pattern, not just this incident. Each bullet: a concrete "
    "broken heuristic the user repeatedly applies. Reference specific qids as evidence.\n\n"
    "  ### Why this question triggered the pattern\n"
    "  2-3 sentences naming the trigger word/structure in this stem that activated the bad "
    "heuristic. Be specific to the stem text.\n\n"
    "  ### What needs to be unlearned\n"
    "  2-3 bullets: the specific belief/shortcut the user must drop before they stop making this "
    "class of mistake. Not 'study more' — name the exact wrong rule running in their head.\n\n"
    "Hard rules:\n"
    "- Pattern-focused, not question-focused. The synthesizer handles the per-question lecture.\n"
    "- If the history is sparse, say so explicitly — don't fabricate patterns.\n"
    "- Senior engineer audience. Be blunt about the dysfunction, not validating."
)

SYSTEM_PROMPT_NUCLEAR_DISTRACTORS = (
    "You are a Distractor Forensics specialist. You receive ONE AWS Advanced Networking question "
    "with all options. Your sole job: forensic-grade autopsy of EACH option — why it exists in "
    "the question writer's mind, what specific misconception it's designed to bait, and exactly "
    "why it fails (or succeeds) under AWS rules.\n\n"
    "Output MARKDOWN only. Structure:\n\n"
    "  ## Distractor forensics — Q<qid>\n\n"
    "  For EACH option (A, B, C, D, E if present):\n\n"
    "  ### Option <X>: <one-line gist>\n"
    "  - **Verdict:** TRAP | CORRECT | NEAR-MISS\n"
    "  - **Why it appears plausible:** the surface logic that makes it look right.\n"
    "  - **The specific misconception it baits:** name the exact wrong belief (e.g. 'TGW route "
    "tables work like VPC route tables', 'NAT GW provides ingress', 'security groups are "
    "stateless'). One precise sentence.\n"
    "  - **Why it actually fails (or works):** the AWS rule that decides. Cite the rule precisely.\n"
    "  - **Forensic tell:** the keyword or phrase in this option that should have triggered "
    "suspicion (e.g. 'the word *transitively* is the tell — peering is non-transitive').\n\n"
    "Hard rules:\n"
    "- Treat every option as evidence in a crime scene. Be precise, not opinionated.\n"
    "- Don't summarize the answer at the end — the synthesizer handles that.\n"
    "- Senior engineer audience. Skip basics."
)

SYSTEM_PROMPT_NUCLEAR_SYNTH = (
    "You are the SYNTHESIZER. You receive: ONE AWS Advanced Networking question the user got "
    "wrong, plus FOUR pieces of prior analysis (boundary spec, pattern analysis, distractor "
    "forensics, optional concept dossier).\n\n"
    "Your job: produce a TIGHT, NON-REDUNDANT lecture on this question. The user has Deep "
    "Review for full mini-class teaching — Nuclear's role is the RULE, the MECHANICS, the "
    "REFLEXES. Not a re-teaching. Read start-to-end like a single coherent argument, no "
    "section repeating another.\n\n"
    "Output MARKDOWN only. EXACTLY 5 sections, in order, each with a unique role. Do not "
    "add headings outside this list. Do not restate the same insight under multiple headings.\n\n"
    "  # Q<qid> — <topic> · NUCLEAR REVIEW\n\n"
    "  ## 0. Question recap\n"
    "  Paraphrase of the stem (5-8 lines) preserving EVERY constraint: services, topology, "
    "scale, latency/throughput numbers, what's blocked or required, what the company is "
    "trying to achieve. Then list EVERY option verbatim, one per line as `- A) <full text>`, "
    "marking the user's pick with `<- your answer` and the correct one with `<- correct`.\n\n"
    "  ## 1. The decisive rule\n"
    "  2-4 sentences max. The ONE rule, mechanic, or boundary that decides this question. "
    "Name it bluntly. If a small ASCII diagram or one-row table makes the rule sharper, "
    "include it — otherwise just prose. No buildup, no 'first let's understand'. Just the "
    "rule.\n\n"
    "  ## 2. Per-option verdict\n"
    "  ONE bullet per option. Format: `- X) <verdict>: <single-sentence reason that ties "
    "back to the decisive rule>`. Verdict ∈ {CORRECT, TRAP, NEAR-MISS}. No paragraphs, no "
    "sub-bullets, no repeating the rule from §1 — just apply it.\n\n"
    "  ## 3. Reasoning chain\n"
    "  4-6 NUMBERED steps. From reading the stem to picking the answer. Each step is one "
    "concrete action (`Step 3: list every destination the workload talks to`). Steps must "
    "compose into a procedure the user can run on a NEW question of the same shape — not a "
    "recap of why this specific qid is what it is.\n\n"
    "  ## 4. Reflexes\n"
    "  3-5 bullets, format: `When you see <stem trigger> → <immediate action>`. Trigger "
    "must be something visible in a stem (a phrase, a service combo, a number range). "
    "Action must be concrete. These are the muscle-memory shortcuts; they should not "
    "duplicate the steps in §3.\n\n"
    "Hard rules:\n"
    "- Use the four input sources as EVIDENCE for §1 and §2. Don't quote them, don't list "
    "them, don't recap them. Integrate.\n"
    "- NO 'What you got wrong' / 'Truth from zero' / 'Common thread' / 'Self-check' / "
    "'Questions in the wild' / 'Mental model to install' sections. Those belong to Deep "
    "Review, not Nuclear.\n"
    "- NO restating §1 inside §2, §3, or §4. The decisive rule is stated ONCE.\n"
    "- Senior engineer audience. Be terse where prose is fine, dense where mechanics matter. "
    "No padding, no hedging, no 'it depends' unless AWS docs are genuinely ambiguous (then "
    "say exactly where).\n"
    "- Respond ENTIRELY in English. No Spanish words, headings, or markers anywhere."
)


def _build_nuclear_msg(label: str, item: Dict[str, Any], extras: Optional[Dict[str, Any]] = None) -> str:
    payload: Dict[str, Any] = {"question": item}
    if extras:
        payload.update(extras)
    return (
        f"{label}\n\n"
        "```json\n"
        + json.dumps(payload, indent=2, ensure_ascii=False)
        + "\n```"
    )


async def _nuclear_fanout(
    item: Dict[str, Any],
    master_excerpts: List[Dict[str, Any]],
    dossier_md: str,
    model: Optional[str],
    provider: str,
    timeout_sec: float,
    progress_cb: Optional[Callable[[str], None]],
) -> Dict[str, str]:
    """Spawn 3 specialist agents in parallel, then run synthesizer with their output."""
    def _emit(s: str) -> None:
        if progress_cb:
            try:
                progress_cb(s)
            except Exception:
                _logger.exception("progress_cb raised")

    boundary_msg = _build_nuclear_msg(
        "Analyze the AWS service boundaries for this question. "
        "Use the concept dossier (if any) as reference material.",
        item,
        {"concept_dossier": dossier_md} if dossier_md else None,
    )
    patterns_msg = _build_nuclear_msg(
        f"Analyze the user's mistake patterns. They have {len(master_excerpts)} prior "
        "wrong-answer reports in their history. Find the pattern, not just this incident.",
        item,
        {"prior_history_excerpts": master_excerpts} if master_excerpts else None,
    )
    distractors_msg = _build_nuclear_msg(
        "Forensic autopsy of every option in this question. Treat each as evidence.",
        item,
    )

    _emit("nuclear: spawning 3 parallel sub-agents (boundary + patterns + distractors)…")
    boundary_t = _run_llm_async(
        SYSTEM_PROMPT_NUCLEAR_BOUNDARY, boundary_msg, model,
        lambda s: _emit(f"boundary: {s}"), timeout_sec,
        provider=provider,
    )
    patterns_t = _run_llm_async(
        SYSTEM_PROMPT_NUCLEAR_PATTERNS, patterns_msg, model,
        lambda s: _emit(f"patterns: {s}"), timeout_sec,
        provider=provider,
    )
    distractors_t = _run_llm_async(
        SYSTEM_PROMPT_NUCLEAR_DISTRACTORS, distractors_msg, model,
        lambda s: _emit(f"distractors: {s}"), timeout_sec,
        provider=provider,
    )
    boundary_md, patterns_md, distractors_md = await asyncio.gather(
        boundary_t, patterns_t, distractors_t,
    )
    _emit("nuclear: 3 sub-agents done. Running synthesizer…")

    synth_extras = {
        "boundary_analysis": boundary_md,
        "pattern_analysis": patterns_md,
        "distractor_forensics": distractors_md,
    }
    if dossier_md:
        synth_extras["concept_dossier"] = dossier_md
    synth_msg = _build_nuclear_msg(
        "Synthesize the four inputs below into the absolute-final-word lecture for this question. "
        "Use ALL of them as evidence; integrate, don't repeat verbatim.",
        item, synth_extras,
    )
    final_md = await _run_llm_async(
        SYSTEM_PROMPT_NUCLEAR_SYNTH, synth_msg, model,
        lambda s: _emit(f"synth: {s}"), timeout_sec,
        provider=provider,
    )
    _emit("nuclear: synthesizer done.")
    return {
        "final": final_md,
        "boundary": boundary_md,
        "patterns": patterns_md,
        "distractors": distractors_md,
    }


def analyze_nuclear_review(
    item: Dict[str, Any],
    master_excerpts: Optional[List[Dict[str, Any]]] = None,
    dossier_md: str = "",
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    timeout_sec: float = 600.0,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, str]:
    """Sync wrapper. Returns dict with keys: final, boundary, patterns, distractors."""
    _logger.info(
        "analyze_nuclear_review: qid=%s, master_excerpts=%d, dossier=%d chars, provider=%s, model=%s",
        item.get("qid"), len(master_excerpts or []), len(dossier_md or ""),
        _normalize_provider(provider), model or "<default>",
    )
    norm = _normalize_item(item)
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        except Exception:
            _logger.exception("failed to set WindowsProactorEventLoopPolicy")
    loop = asyncio.new_event_loop()
    try:
        asyncio.set_event_loop(loop)
        return loop.run_until_complete(_nuclear_fanout(
            norm, master_excerpts or [], dossier_md or "",
            model, provider, timeout_sec, progress_cb,
        ))
    finally:
        try:
            pending = asyncio.all_tasks(loop)
            for t in pending:
                t.cancel()
            if pending:
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        except Exception:
            _logger.exception("error draining loop tasks")
        loop.close()


# ============================================================================
# God Synthesizer — meta-coach with cross-layer awareness.
# Reads ai_coach_reports.md raw + COMPRESSED digests of deep + nuclear (one
# line per qid) + stats sidecar. Avoids the Lovecraftian-monster anti-pattern
# of swallowing every byte: digests carry the SIGNAL, not the lecture.
# ============================================================================

SYSTEM_PROMPT_GOD = (
    "You are the GOD SYNTHESIZER. The user has been studying AWS Advanced Networking "
    "(ANS-C01) using a 3-layer system: AI Coach (cheap diagnostic briefs), Deep Review "
    "(per-question full mini-classes), and Nuclear (multi-agent ultradios for critical "
    "questions). Your job: read across ALL THREE layers and produce a unified study plan "
    "that no single layer could.\n\n"
    "You receive:\n"
    "1. `ai_coach_reports.md` (raw): brief diagnostics per batch.\n"
    "2. Deep digest: one line per qid with topic + the canonical takeaway from that qid's "
    "Deep Review. NOT the full lecture — the signal.\n"
    "3. Nuclear digest: one line per qid with topic + the decisive rule from that qid's "
    "Nuclear Review.\n"
    "4. Stats sidecar (JSON): hot zone, top reincident qids, totals.\n\n"
    "Output MARKDOWN only. EXACTLY these sections, in order:\n\n"
    "  # God Synthesizer — Cross-Layer Study Plan\n\n"
    "  ## 1. Where you actually are\n"
    "  3-5 sentences. State the user's current learning status with surgical honesty: "
    "what concepts are LOCKED IN (they have Deep + Nuclear and stats show no recent fail), "
    "what's IN PROGRESS (covered by Deep but still failing per stats), what's UNTOUCHED "
    "(qids in stats with no Deep/Nuclear yet). Use the digests to verify, not guess.\n\n"
    "  ## 2. Recurring failure patterns\n"
    "  3-5 bullets. Cross-cutting failure shapes that show up across MULTIPLE qids — "
    "patterns visible only because you're seeing all 3 layers. Format: `**<pattern name>**: "
    "<one-line description>. Evidence: Q<n>, Q<m>, Q<p>.` Cite the qids that prove the "
    "pattern. Do NOT invent patterns that aren't supported by 2+ qids.\n\n"
    "  ## 3. The next 5 actions, ranked\n"
    "  Numbered 1-5. Each item: a CONCRETE action with a CONCRETE target. Format: "
    "`1. <Action> — Q<n>, Q<m>... (Reason: <one line>).` Examples of valid actions: "
    "'Run Nuclear on Q<n>', 'Re-do Deep on Q<n> with focus on <subtopic>', "
    "'Drill these 5 qids in a custom round', 'Build a concept dossier for <topic>'. "
    "Rank by ROI: what fixes the most failure surface for the least token cost.\n\n"
    "  ## 4. What NOT to do\n"
    "  2-3 bullets. Common time-sinks for THIS user given their state. e.g. 'Don't run "
    "Nuclear on Q<n> — Deep already covered the rule, the failure is execution speed', "
    "or 'Don't add more AI Coach passes on already-Deep-reviewed qids — diminishing "
    "returns'. Be specific.\n\n"
    "Hard rules:\n"
    "- Cite qid numbers as evidence. Patterns without qid evidence are speculation.\n"
    "- Don't repeat what any single layer already said. You are the layer ABOVE them.\n"
    "- Stats sidecar is ground truth for 'is the user still failing this?' Trust it over "
    "any narrative in the briefs.\n"
    "- Length: tight. ~400-700 words total. This is a battle plan, not an essay."
)


# Capture group 1 = qid, group 2 = topic (after the em-dash, until newline).
_DEEP_HEADING_RE = re.compile(
    r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)\b[^>]*-->\s*\n"
    r"#\s+[^\n]*?\n+"  # the date heading
    r"###\s+Q\d+\s*[—\-]\s*([^\n]+)",
    re.MULTILINE,
)

_NUCLEAR_HEADING_RE = re.compile(
    r"<!--\s*NUCLEAR_REVIEW\s+qid=(\d+)\b[^>]*-->\s*\n"
    r"#\s+[^\n]*?\n+"  # date heading
    r"##\s+Final synthesis\s*\n+"
    r"#\s+Q\d+\s*[—\-]\s*([^·\n]+?)(?:·|\n)",
    re.MULTILINE,
)


def _extract_deep_digest(deep_md: str) -> str:
    """One-line digest per qid from deep_review_reports.md.

    Picks the topic from `### Q<n> — <topic>` and the takeaway from the first
    bullet of `## Common thread` (Deep Review's own self-summary). Falls back
    to the first bullet of section "1. Why you screwed it up" if Common thread
    is missing. Pure local parse, zero tokens.
    """
    if not deep_md or not deep_md.strip():
        return "_(no Deep Review history yet)_"

    # Split file into per-qid sections by the qid marker.
    # Each section starts at <!-- DEEP_REVIEW_QID qid=N ... -->.
    sections = re.split(r"(?=<!--\s*DEEP_REVIEW_QID\s+qid=\d+)", deep_md)
    lines: List[str] = []
    seen: Dict[int, str] = {}  # qid -> latest digest (so re-runs overwrite older)
    for sec in sections:
        m = re.match(r"<!--\s*DEEP_REVIEW_QID\s+qid=(\d+)", sec)
        if not m:
            continue
        qid = int(m.group(1))
        topic_m = re.search(r"###\s+Q\d+\s*[—\-]\s*([^\n]+)", sec)
        topic = (topic_m.group(1).strip() if topic_m else "").strip("· ").strip()

        # Try Common thread → first bullet
        ct_m = re.search(
            r"##\s+Common thread\s*\n+\s*-\s*([^\n]+)",
            sec, re.IGNORECASE,
        )
        # Fallback: section 1 "Why you screwed it up" first sentence
        if not ct_m:
            ct_m = re.search(
                r"\*\*1\.\s*Why you screwed it up\*\*\s*\n+([^\n]+)",
                sec, re.IGNORECASE,
            )
        takeaway = (ct_m.group(1).strip() if ct_m else "").strip()
        # Truncate aggressively — this is a one-liner.
        if len(takeaway) > 220:
            takeaway = takeaway[:217].rstrip() + "…"
        if not takeaway:
            takeaway = "(no takeaway extracted)"

        seen[qid] = f"- Q{qid} [{topic or '?'}]: {takeaway}"

    if not seen:
        return "_(deep_review_reports.md present but no qid sections parsed)_"
    for qid in sorted(seen.keys()):
        lines.append(seen[qid])
    return "\n".join(lines)


def _extract_nuclear_digest(nuclear_md: str) -> str:
    """One-line digest per qid from nuclear_reports.md.

    Picks the topic from the `# Q<n> — <topic> · NUCLEAR REVIEW` heading and
    the first paragraph of `## 1. The decisive rule` (new prompt) or the
    `**The single sentence...**` line (old prompt) as the takeaway.
    """
    if not nuclear_md or not nuclear_md.strip():
        return "_(no Nuclear Review history yet)_"

    sections = re.split(r"(?=<!--\s*NUCLEAR_REVIEW\s+qid=\d+)", nuclear_md)
    seen: Dict[int, str] = {}
    for sec in sections:
        m = re.match(r"<!--\s*NUCLEAR_REVIEW\s+qid=(\d+)", sec)
        if not m:
            continue
        qid = int(m.group(1))
        topic_m = re.search(
            r"#\s+Q\d+\s*[—\-]\s*(.+?)(?:·\s*NUCLEAR REVIEW|\n)",
            sec,
        )
        topic = (topic_m.group(1).strip() if topic_m else "").strip()

        # New prompt: ## 1. The decisive rule  → next non-empty paragraph
        rule_m = re.search(
            r"##\s+1\.\s*The decisive rule\s*\n+([^\n][^\n]*(?:\n[^\n#]+)*)",
            sec, re.IGNORECASE,
        )
        # Old prompt fallback: **The single sentence... :** <text>
        if not rule_m:
            rule_m = re.search(
                r"\*\*The single sentence[^*]*\*\*\s*([^\n]+)",
                sec, re.IGNORECASE,
            )
        takeaway = (rule_m.group(1).strip() if rule_m else "").strip()
        # Collapse internal whitespace, truncate.
        takeaway = re.sub(r"\s+", " ", takeaway)
        if len(takeaway) > 260:
            takeaway = takeaway[:257].rstrip() + "…"
        if not takeaway:
            takeaway = "(no decisive rule extracted)"

        seen[qid] = f"- Q{qid} [{topic or '?'}]: {takeaway}"

    if not seen:
        return "_(nuclear_reports.md present but no qid sections parsed)_"
    return "\n".join(seen[q] for q in sorted(seen.keys()))


def analyze_god_synthesizer(
    ai_coach_md: str,
    deep_md: str,
    nuclear_md: str,
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    max_chars_ai_coach: int = 180_000,
    timeout_sec: float = 600.0,
    progress_cb: Optional[Callable[[str], None]] = None,
    stats_payload: Optional[Dict[str, Any]] = None,
) -> str:
    """Cross-layer synthesizer: AI Coach raw + Deep digest + Nuclear digest + stats.

    Digests keep token cost ~1.5x of plain Meta-Coach (not 3x), while giving the
    model real visibility into the user's full study state across all 3 layers.
    """
    _logger.info(
        "analyze_god_synthesizer: ai_coach_chars=%d, deep_chars=%d, nuclear_chars=%d, provider=%s, model=%s",
        len(ai_coach_md or ""), len(deep_md or ""), len(nuclear_md or ""),
        _normalize_provider(provider), model or "<default>",
    )

    ai_coach_text = (ai_coach_md or "").strip()
    truncated_note = ""
    if len(ai_coach_text) > max_chars_ai_coach:
        ai_coach_text = ai_coach_text[-max_chars_ai_coach:]
        marker = ai_coach_text.find("<!-- AI_COACH_BATCH")
        if marker > 0:
            ai_coach_text = ai_coach_text[marker:]
        truncated_note = (
            f"\n_(NOTE: ai_coach_reports.md was {len(ai_coach_md)} chars; trimmed oldest "
            f"to fit {max_chars_ai_coach}.)_\n"
        )

    deep_digest = _extract_deep_digest(deep_md or "")
    nuclear_digest = _extract_nuclear_digest(nuclear_md or "")

    stats_block = ""
    if stats_payload:
        stats_block = (
            "Stats sidecar (ground truth for 'is user still failing this?'):\n"
            "```json\n"
            + json.dumps(stats_payload, indent=2, ensure_ascii=False)
            + "\n```\n\n"
        )

    user_msg = (
        "You have all 3 layers of the user's study history below. Synthesize per the rules.\n\n"
        f"{stats_block}"
        f"{truncated_note}"
        "## Layer 1 — AI Coach reports (raw)\n"
        "```markdown\n"
        f"{ai_coach_text or '_(empty)_'}\n"
        "```\n\n"
        "## Layer 2 — Deep Review digest (one line per qid, topic + canonical takeaway)\n"
        f"{deep_digest}\n\n"
        "## Layer 3 — Nuclear Review digest (one line per qid, topic + decisive rule)\n"
        f"{nuclear_digest}\n"
    )
    return _run_llm_sync(
        SYSTEM_PROMPT_GOD, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )


# ============================================================================
# Diagram — produce a self-contained HTML study artifact for ONE question.
# Reads any prior artifacts (deep/nuclear/coach excerpts) to avoid re-deriving
# from scratch. Output is RAW HTML (no markdown fences) so the GUI can save it
# as Q<N>_<topic>_diagram.html and open it in the default browser.
# ============================================================================

SYSTEM_PROMPT_DIAGRAM = (
    "You are a senior AWS networking instructor producing a visual study artifact for a single "
    "ANS-C01 question. Audience: a senior network engineer studying the AWS exam; assume strong "
    "networking fundamentals, but keep the artifact AWS-first and vendor-neutral. Aesthetic target: high-production single-page study card — "
    "ONE dramatic hero diagram + supporting cards, NOT a wall of small isolated diagrams.\n\n"
    "Output: RAW HTML5 only. No markdown fences, no preamble, no '```html', no commentary — the "
    "first character of your output must be '<' (e.g., `<!doctype html>`). Inline CSS only. Dark "
    "theme. Self-contained file (no external assets, no JS, no CDNs).\n\n"
    "=== REQUIRED <head> + <style> SCAFFOLD (use this palette and style verbatim) ===\n"
    "  :root {\n"
    "    --bg:#081321; --panel:#0d1f35; --panel2:#0b1829;\n"
    "    --cyan:#4dd0ff; --green:#51e28a; --yellow:#ffd54f; --orange:#ffb74d;\n"
    "    --red:#ff6b6b; --purple:#aa8cff; --text:#e6eef8; --muted:#96a9bf;\n"
    "    --border:#1a3654;\n"
    "  }\n"
    "  * { box-sizing:border-box; margin:0; padding:0; }\n"
    "  body { font-family:'Segoe UI', system-ui, sans-serif; "
    "background:radial-gradient(circle at top, #10233d 0%, var(--bg) 48%); color:var(--text); "
    "padding:24px; }\n"
    "  .wrap { max-width:1440px; margin:0 auto; }\n"
    "  h1 { color:var(--cyan); font-size:1.45rem; margin-bottom:4px; }\n"
    "  .sub { color:var(--muted); font-size:.88rem; margin-bottom:14px; }\n"
    "  .answer { display:inline-block; background:#0e2a1d; border:1px solid #24583a; "
    "color:var(--green); padding:10px 14px; border-radius:12px; font-weight:800; "
    "margin-bottom:18px; font-size:.95rem; }\n"
    "  .card { background:linear-gradient(180deg, var(--panel), var(--panel2)); "
    "border:1px solid var(--border); border-radius:18px; padding:18px; "
    "box-shadow:0 8px 24px rgba(0,0,0,.22); margin-bottom:18px; }\n"
    "  .card h2 { color:var(--cyan); font-size:1.05rem; margin-bottom:10px; }\n"
    "  .grid-cards { display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); "
    "gap:14px; margin-top:14px; }\n"
    "  .study-rail { margin:18px 0 20px; padding:6px 2px 2px; }\n"
    "  .study-rail h2 { color:var(--cyan); font-size:1.05rem; margin-bottom:8px; }\n"
    "  .study-row { display:grid; grid-template-columns:minmax(180px,260px) 1fr; gap:18px; "
    "align-items:start; padding:13px 0; border-top:1px solid rgba(150,169,191,.22); }\n"
    "  .study-row:last-child { border-bottom:1px solid rgba(150,169,191,.22); }\n"
    "  .study-key { font-weight:900; font-size:.9rem; line-height:1.35; }\n"
    "  .study-key.good { color:var(--green); } .study-key.bad { color:var(--red); } "
    ".study-key.warn { color:var(--yellow); }\n"
    "  .study-row p { color:#c8d5e6; font-size:.88rem; line-height:1.6; }\n"
    "  .study-row p b { color:var(--text); }\n"
    "  .opt { border-radius:14px; padding:12px; border:1px solid var(--border); }\n"
    "  .opt.bad { background:#1a0d0d; border-color:#3b2a46; }\n"
    "  .opt.good { background:#0d1f0d; border-color:#24583a; }\n"
    "  .opt .t { font-weight:800; margin-bottom:6px; font-size:.9rem; }\n"
    "  .opt .t.bad { color:var(--red); } .opt .t.good { color:var(--green); }\n"
    "  .opt p { color:#c8d5e6; font-size:.85rem; line-height:1.55; }\n"
    "  table { border-collapse:collapse; width:100%; margin:10px 0; }\n"
    "  th,td { border:1px solid var(--border); padding:8px 12px; font-size:.85rem; text-align:left; }\n"
    "  th { background:#0d2137; color:var(--cyan); }\n"
    "  td.yes { color:var(--green); font-weight:700; } td.no { color:var(--red); font-weight:700; }\n"
    "  .key { color:var(--yellow); font-weight:700; }\n"
    "  .memorize { background:#1a1810; border:1px solid #5b4a1f; border-radius:14px; "
    "padding:14px; color:var(--yellow); font-weight:700; font-size:1rem; line-height:1.5; "
    "margin-top:14px; }\n"
    "  ul { padding-left:20px; } li { margin:5px 0; font-size:.88rem; line-height:1.55; "
    "color:#c8d5e6; }\n"
    "  svg { display:block; margin:4px auto; max-width:100%; height:auto; }\n"
    "  svg text { font-family:'Segoe UI', system-ui, sans-serif; }\n"
    "  svg path[stroke]:not([fill]), svg polyline[stroke]:not([fill]) { fill:none; }\n\n"
    "=== AUDIENCE-AWARE WRITING (the artifact must TEACH, not confirm) ===\n"
    "The reader understands BGP, routing tables, RIB selection, DNS, VPNs, firewalls, NAT, "
    "MTU, and failover mechanics. He does NOT automatically know AWS-specific terms, limits, "
    "or service behaviors. Default failure mode of these "
    "artifacts: they read like exam answer keys for someone who already mastered ~75% of the "
    "AWS material — they confirm rather than explain.\n\n"
    "Hard rules to break that pattern:\n"
    "  1. INLINE GLOSS on first use of any AWS-specific term per artifact. Don't say "
    "'apex / zone apex' alone — say 'apex (the bare domain like example.com, no subdomain "
    "prefix; CNAME is illegal there per DNS RFC)'. Don't say 'LBR' alone — say 'LBR (Latency-"
    "Based Routing — Route 53 picks the AWS Region with lowest measured latency from the "
    "resolver's IP)'. Don't say 'AWS backbone' alone — say 'AWS backbone (the private fiber "
    "network connecting all AWS Regions and edge PoPs; bypasses public Internet hops)'.\n"
    "  2. NETWORK PRIMITIVE, not vendor analogy. Explain AWS behavior through exam-relevant "
    "primitives such as route-table lookup, longest-prefix match, BGP route attributes, DNS "
    "TTL/caching, health checks, NAT state, symmetric return path, and service quotas. Do NOT "
    "use Cisco/Juniper/vendor CLI framing, MPLS-L3VPN analogies, VRF-per-VPC analogies, "
    "route-target language, or RFC detours unless the question explicitly asks for them.\n"
    "  3. NAME THE MECHANISM, never just the outcome. 'Sub-second failover' is an outcome — "
    "the mechanism is 'GA endpoint health-check withdraws the unhealthy endpoint group from "
    "the anycast advertisement at the edge'. Always state the mechanism alongside the outcome.\n"
    "  4. EXPLAIN THE TRAP for every wrong option (see Option Read-Through structure). The trap "
    "is what teaches the discrimination skill the exam tests; the right answer alone teaches "
    "nothing about why the wrong ones tempted.\n\n"
    "=== REQUIRED DOCUMENT STRUCTURE (in order) ===\n"
    "  1. <h1>Q<N> — <topic></h1>\n"
    "  2. <div class='sub'>one-line stem recap (~120 chars)</div>\n"
    "  3. <div class='answer'>✅ Correct: <letter(s)> — <short reason></div>\n"
    "  4. HERO ARCHITECTURE DIAGRAM in a .card (see HERO STYLE below). This is THE main visual; "
    "spend the most effort here.\n"
    "  5. AWS TERM QUICK REFERENCE rail BEFORE the option read-through. "
    "<div class='study-rail'><h2>AWS Terms in Play — quick gloss</h2>... 4-8 .study-row, one "
    "per AWS-specific term that the option read-through will use (e.g. 'anycast IP', 'AWS "
    "backbone', 'apex / zone apex', 'alias record', 'LBR (latency-based routing)', 'GA endpoint "
    "group', 'TLS pass-through', 'health-check failover'). Left = the term (study-key warn). "
    "Right = ONE clear sentence: literal AWS definition + why it matters to this question "
    "(e.g. 'anycast IP — one static IP advertised from many AWS edge locations; clients enter "
    "the AWS edge nearest to them and AWS forwards over its backbone.'). This panel is "
    "MANDATORY whenever the question uses AWS-specific networking terms (almost every question). "
    "Skip only for pure config-flag questions with no jargon.\n"
    "  6. Option read-through rail, NOT a wall of concept cards: "
    "<div class='study-rail'><h2>Option Read-Through — where each answer lands</h2>... "
    "Use one .study-row per exam option. Left side = <div class='study-key good|bad|warn'>"
    "✓/✗ A — short title</div>. Right side structure:\n"
    "    For WRONG options: TWO short paragraphs. First (italic, muted color #96a9bf) is "
    "'Why it looks right at first glance' — articulate the trap honestly, the brain reflex that "
    "would lead a network engineer or rushed exam taker to pick this. Second (normal color) is 'Why it actually fails' "
    "— the AWS-specific fact that disqualifies it. This trap/truth structure is what teaches; "
    "skipping the 'looks right' half makes the option a checklist confirmation, not a lesson.\n"
    "    For CORRECT options: one paragraph explaining what makes it the win, with the key "
    "AWS mechanism named explicitly (not just 'sub-second failover' but 'GA endpoint health "
    "checks → BGP-style withdraw at edge PoP, sub-second').\n"
    "    Mark the user's pick inline in the left label ('your pick ✓' or 'your pick ✗').\n"
    "  6. FLOW SEQUENCE section(s) — at least 2 decision-flow SVGs covering the resolution "
    "flow (see FLOW SEQUENCE STYLE below). These complement the hero diagram by showing the "
    "lookup/decision first, then the selected path, then the fallback behavior. Required when "
    "the question involves any data flow, routing decision, or stateful processing. Skip ONLY "
    "if the question is purely about a config flag with no flow (rare).\n"
    "  6a. MASTERCLASS panel (.card with h2 'Why this design — architectural reasoning'). "
    "Three .study-rail rows MANDATORY:\n"
    "    • 'Why this design exists' — historical/architectural reason the AWS service was built "
    "this way. E.g., 'GA exists because Route 53 LBR can only steer DNS resolution; once the "
    "client has the IP, it traverses random ISP paths. AWS needed an L4 anycast layer to also "
    "control the data-plane path, not just the destination choice; DNS picks an endpoint, while "
    "Global Accelerator also changes where traffic enters and crosses the AWS backbone.'\n"
    "    • 'Common misconceptions' — 3-5 bullet list of things engineers wrongly assume. Pull "
    "these from prior_artifacts (deep_review_section, nuclear_section) when present. E.g., 'GA "
    "terminates TLS' (no), 'GA needs CloudFront in front' (no), 'one ACM cert covers all Regions' "
    "(no, that's CloudFront only), 'LBR + GA are redundant' (no, LBR steers DNS, GA steers data).\n"
    "    • 'Concrete numbers and tradeoffs' — real magnitudes the engineer should remember. "
    "E.g., 'GA reduces p99 cross-Region latency 30-60% typical', 'health-check failover ~1s vs "
    "DNS TTL 60s+ for LBR', '2 static anycast IPs per accelerator', 'standard accelerator costs "
    "$0.025/hr + $0.015/GB outbound vs LBR essentially free'. If exact numbers aren't known, "
    "give order of magnitude and label as such.\n"
    "  7. .card with a compatibility/comparison table when relevant (e.g. 'SiteLink VIF support', "
    "'TGW vs PrivateLink', 'route precedence'). Use td.yes / td.no classes.\n"
    "  8. .study-rail 'Integrated Mental Model' when the concept needs explanation. Use rows like "
    "'Exam mental model', 'Why X beats Y', 'Failure reflex'. Avoid separate standalone concept boxes.\n"
    "  9. .memorize callout at the very end: ONE bold sentence the user can memorize for the exam.\n\n"
    "=== HERO ARCHITECTURE STYLE (the ONE big SVG at top) ===\n"
    "Single SVG. viewBox starts at '0 0 1200 640' minimum; if topology is complex, grow to "
    "1400×800. Empty space between any two component boxes ≥ 40px. NEVER compress — grow the "
    "viewBox instead. The SVG must breathe.\n\n"
    "Required <defs>:\n"
    "  • Arrow markers, one per accent color: 'ag' #51e28a (green), 'ar' #ff6b6b (red), "
    "'ay' #ffd54f (yellow), 'ac' #4dd0ff (cyan), 'ao' #ffb74d (orange), 'ap' #aa8cff (purple).\n"
    "  • Glow filter:  <filter id='glow'><feGaussianBlur stdDeviation='4' result='b'/>"
    "<feMerge><feMergeNode in='b'/><feMergeNode in='SourceGraphic'/></feMerge></filter>\n"
    "    Use ONCE on THE pivotal component the answer hinges on (DX-GW, TGW, ANFW endpoint, etc).\n\n"
    "Layout zones (top to bottom, no overlap):\n"
    "  • y=0–80   BEFORE/AFTER band. Keep each line <= 90 visible characters. If it would be "
    "longer, split into two <tspan> lines or shorten it. Two centered <text>: y=22 muted "
    "'#78909c' problem state, y=48 green '#51e28a' what changes. NOTHING ELSE in this band.\n"
    "  • y=110+   Topology zones and components. Any dashed callout/header such as 'Weekend "
    "bridge' must have its text INSIDE its dashed rectangle and below y=95.\n"
    "  • bottom   Answer strip: draw ONE box per answer option (usually 4–6 boxes). Wrong: "
    "fill #1a0d0d stroke #ef5350. Correct: fill #0d1f0d stroke #69f0ae sw=2. ≥30px padding "
    "above strip, no intrusion.\n\n"
    "Zone rects (rounded rx=12, fill='none', stroke 2px):\n"
    "  on-prem DC = purple #7c4dff,  VPC/region = cyan #1565c0,  AZ = teal #00A4A6 DASHED.\n"
    "  Zone header bold 12px tinted to stroke color. Do NOT use flag emoji in SVG zone headers; "
    "some browsers render them as stray country-code text like 'us'. Use plain labels such as "
    "'ON-PREM DC-EAST' or 'AWS Region us-east-1'.\n"
    "  ⚠ Sanity check: an on-prem element (Local ISP, customer router, on-prem DC) MUST live "
    "inside an on-prem zone, NEVER inside a VPC zone. Verify each component's (x,y) sits inside "
    "the right zone before emitting.\n\n"
    "Component boxes (rounded rx=8–10, service-tinted dark fill, colored stroke):\n"
    "  TILE CONTENT BUDGET — identity only. Line 1 BOLD service name (≥14px), line 2 "
    "instance/role (≤25 chars, ≥12px). NO third line of detail inside the tile. Cert names, "
    "CIDRs, AZ IDs, ASN numbers, beta flags, route-table IDs all go in a NUMBERED LEGEND "
    "table inside the same .card below the SVG. If a tile feels like it needs more text, "
    "it's the wrong abstraction — push the detail to the legend and let the tile breathe.\n\n"
    "REPEATED IDENTICAL COMPONENTS. When the topology contains N≥4 instances of the same "
    "component (8 ALBs, 6 VPCs, 4 AZs, multiple Direct Connect locations), draw ONE fully "
    "detailed instance and collapse the rest:\n"
    "  • 1 detailed tile in front (normal stroke, full identity).\n"
    "  • 2 ghost tiles behind it (offset translate(+6,+6) and (+12,+12), 50% opacity, no text).\n"
    "  • Badge '×N' in the top-right corner of the front tile (yellow background, bold).\n"
    "  • Do NOT draw N parallel arrows from a fan-out point — draw ONE thick arrow with label "
    "'→ nearest of N (anycast / latency / health)'. Per-instance variation goes in the legend.\n\n"
    "NUMBERED CANONICAL PATH. Place ① ② ③ ... circles (r=12, fill=cyan #4dd0ff, white text, "
    "bold) directly ON the canonical happy-path arrow chain at each hop. Other arrows stay "
    "unnumbered. Below the SVG inside the same .card, add a numbered legend list:\n"
    "  ① <hop event> — <state change at this hop>\n"
    "  ② <hop event> — <state change at this hop>\n"
    "The numbers anchor the eye on what to read first; the legend is where the actual "
    "per-hop story lives.\n\n"
    "Magic path overlay: when ONE non-obvious feature is the answer (SiteLink, GWLB hairpin, "
    "edge-association RT), wrap that region in a yellow dashed rect "
    "(stroke='#ffd740' stroke-dasharray='8,4') with '⚡ <feature>' bold-yellow header inside.\n\n"
    "Arrows — semantic color, stroke-width 2–2.5, ORTHOGONAL ONLY (horizontal or vertical "
    "segments; use waypoints, never diagonals):\n"
    "  green = correct forward path, cyan = generic flow, yellow = magic answer flow, "
    "red = blocked/dropped, orange dashed = physical L1 (DX/fiber), purple dashed = on-prem link.\n"
    "  Each arrow gets one short label (≤30 chars) outside the path, never on top of a box.\n"
    "  Arrowhead marker color MUST match the line stroke: green line uses green marker, yellow "
    "line uses yellow marker, red line uses red marker. Never put a red marker on a yellow/green "
    "connector; it makes the visual say 'drop' when the text says 'reroute'.\n"
    "  Do NOT draw freehand-looking arcs, circles, or sweeping curved callouts around components. "
    "For emphasis, use a dashed rect/callout; for movement, use a clean center-edge connector.\n"
    "  Every connector <path> or <polyline> MUST include fill='none'. Browser default fill is "
    "black and turns open connector paths into giant black polygons. This is a hard rule.\n\n"
    "=== FLOW SEQUENCE PANELS (after the hero, ≥2 of them) ===\n"
    "Prefer route/decision flow panels over UML lifelines. Each flow SVG should read like a "
    "networker's troubleshooting whiteboard: 'what route/condition won?' then 'where does the "
    "packet go?' then 'what changes during failure?'. Use vertical lifelines ONLY when the "
    "question is truly about request/response timing.\n\n"
    "DEPTH REQUIREMENT — the #1 failure mode for these panels is becoming '5 boxes with arrows "
    "between them' that says nothing a sentence wouldn't. To prevent that, every flow panel "
    "MUST show STATE CHANGE at each hop, not just transit:\n"
    "  • Routing/forwarding panel: at each hop annotate src/dst IP, header rewrites, NAT or "
    "encapsulation only when explicitly relevant to the AWS service, TTL behavior, MTU, "
    "route-table consulted, and which entry won (LPM / longest match / propagated vs static). "
    "Do not invent backbone encapsulation details.\n"
    "  • Connection/handshake panel: at each lane annotate what the component DOES with the "
    "packet (accepts SYN, terminates TLS with cert X, reads SNI, rewrites Host header, decrypts, "
    "re-encrypts to backend, etc). 'Forwards' is not a state change — say what changed.\n"
    "  • DNS/resolver panel: show the recursive chain — which resolver/PHZ/endpoint matched, "
    "which CNAME/alias was followed, what the cached TTL is, and what was returned to the client.\n"
    "  • Security/firewall panel: show 5-tuple BEFORE inspection, the rule/stateful entry that "
    "matched, and the 5-tuple AFTER (post-NAT, post-rewrite). Mark return path symmetry "
    "explicitly (✓ symmetric / ✗ asymmetric → drop).\n"
    "If a flow panel reads 'Browser → DNS → GA → ALB → Target' with no extra annotations, it is "
    "automatically failing this rule. Add the per-hop state column.\n\n"
    "Hard rules:\n"
    "  • Use 2-3 large panels: (1) lookup/decision card, (2) selected path chain WITH per-hop "
    "state annotations (above), (3) failure/fallback recompute. Panel (3) is MANDATORY whenever "
    "the question involves health checks, failover, route withdrawal, or HA topology — do not "
    "skip it just because it's tedious.\n"
    "  • The lookup card must explicitly show WINNER vs fallback/loser routes or conditions. "
    "Use green/yellow for winner, muted gray for installed-but-not-used fallback, red for invalid.\n"
    "  • The selected path chain uses component boxes connected left→right with orthogonal "
    "arrows. Each node label is short: service/name on line 1, role on line 2. PER-HOP STATE "
    "(see DEPTH REQUIREMENT) goes in a small <text> column directly under each node, NOT in the "
    "node itself.\n"
    "  • If failure behavior matters, draw a separate bottom panel with the withdrawn route as "
    "an actual <path> arrow in red dashed style (stroke='#ff6b6b' stroke-dasharray='6,4') AND "
    "the replacement route as an actual green <path> arrow. Both arrows MUST be drawn — do not "
    "describe failover in text without showing the arrows. Label the trigger ON the red arrow "
    "(e.g. 'eu-west-1 ALB unhealthy → withdraw') and the recovery ON the green arrow (e.g. "
    "'reroute to nearest healthy: eu-central-1, sub-second'). Self-check: if the panel has "
    "≤2 arrows total or the words 'failover'/'unhealthy'/'withdrawn' appear only in <text> "
    "without a matching dashed-red <path>, the panel is failing this rule and must be redrawn.\n"
    "    Failure panel layout: left-to-right chain only: failed/withdrawn object → recompute/"
    "decision chip → replacement healthy object. Put loser comparisons in a separate side box; "
    "do not route arrows through or behind that box.\n"
    "  • Every arrow label must be outside the line and ≤40 chars. Put longer explanation in "
    "nearby <text> rows or a normal HTML .card below the SVG.\n"
    "  • Every connector <path> or <polyline> MUST include fill='none'. No exceptions.\n"
    "  • Avoid diagonal lifeline-style arrows unless they are genuinely the clearest format; "
    "for routing, BGP, DNS, firewall, and TGW questions, use decision-flow panels.\n\n"
    "=== FLOW RECIPE LIBRARY (choose the matching recipe) ===\n"
    "Routing / BGP / TGW / DX / VPN questions:\n"
    "  A. Draw a RIB/route-table decision card first. Rows: WINNER, fallback, rejected/missing. "
    "Include prefix length, target, and why it won (LPM, propagation, association, priority).\n"
    "  B. Draw the forwarding chain second: source → local edge → AWS construct → destination. "
    "Under each hop, annotate src/dst IP, encapsulation, route-table consulted, winning entry. "
    "Do not show a path unless the decision card already explained why it was selected.\n"
    "  C. If failover matters, draw a recompute panel: red dashed withdrawn route, green "
    "replacement route, one-line trigger (BGP withdraw, health check, route propagation change), "
    "and convergence time-order (sub-second / TTL-bound / minutes).\n"
    "Direct Connect / DXGW / TGW / VIF relationship labels (mandatory when these terms appear):\n"
    "  A. Label the left-side VIF relationship precisely. A private VIF or transit VIF is created "
    "on a DX connection and has a target/attachment to a gateway. Use labels such as "
    "'private VIF target/attachment to DXGW' or 'transit VIF attachment/target to DXGW'. "
    "Do NOT call this a TGW association.\n"
    "  B. Label the right-side DXGW relationship with AWS's association term. Use "
    "'DXGW association to VGW', 'DXGW association to TGW', or 'DXGW association to Cloud WAN "
    "core network segment' as appropriate. If the diagram shows multiple TGWs, put a visible "
    "label on the fan-out arrows: 'DXGW associations to TGWs'.\n"
    "  C. Label VPC-side relationships separately: 'TGW VPC attachment', 'TGW VPN attachment', "
    "or 'VGW attached to VPC'. This avoids collapsing VIF targets, DXGW associations, and TGW "
    "attachments into one vague arrow.\n"
    "  D. Transit VIF diagrams must show: DX connection -> transit VIF -> DXGW -> DXGW association "
    "to TGW(s) or Cloud WAN. Never draw or state 'transit VIF terminates directly on TGW/VGW/VPC'.\n"
    "  E. Private VIF diagrams must show: DX connection -> private VIF -> VGW OR DXGW -> "
    "DXGW associations to VGWs. Do not mix TGWs and VGWs on the same DXGW unless the question "
    "explicitly describes separate DXGWs.\n"
    "  F. When quotas matter, use current exam-safe numbers in captions/tables: dedicated DX "
    "supports up to 4 transit VIFs per connection and up to 50 private/public VIFs per "
    "connection; hosted DX supports one VIF total; a DXGW supports up to 6 TGWs, up to 20 VGWs, "
    "up to 30 private/transit VIFs, and up to 200 DXGWs per account. If uncertain, phrase as "
    "'current AWS quota' and make the topology decision hinge on gateway pairing/cost, not a "
    "stale quota gotcha.\n"
    "DNS / Resolver / Private hosted zone questions:\n"
    "  A. Draw query origin and resolver decision card first: which resolver rule / PHZ / endpoint "
    "matched, what TTL, and which candidate was ignored.\n"
    "  B. Then draw query path and response path separately if they differ. Annotate each hop "
    "with the record type returned (A / AAAA / CNAME / alias) and cache state.\n"
    "Security / inspection / firewall questions:\n"
    "  A. Draw policy/route decision first: which route table or rule sends traffic to inspection.\n"
    "  B. Then draw forward path AND return path on the SAME panel; annotate 5-tuple before/after "
    "any NAT, and explicitly mark symmetric vs asymmetric return (asymmetric = stateful drop).\n"
    "PrivateLink / endpoints / service access questions:\n"
    "  A. Draw name-resolution or route-selection first, then endpoint ENI/service path.\n"
    "  B. Add a small 'not public internet' or 'not transitively reachable' note when that is the trap.\n"
    "Handshake / lifecycle / stateful flows (TLS, TCP, BGP session establishment, DHCP, OAuth):\n"
    "  A. Use VERTICAL sequence-diagram with 3-5 swim lanes (e.g. Browser | Edge | LB | Target). "
    "Time flows TOP to BOTTOM. Lane headers are bold, vertical lifelines dashed gray.\n"
    "  B. Each message is a horizontal arrow between two lanes with a SHORT label (SYN, "
    "ClientHello+SNI, ServerHello+cert, Finished, GET, 200 OK). Group related exchanges with a "
    "translucent rounded rectangle (TLS handshake / HTTP request / etc).\n"
    "  C. Below the lifeline, add a small annotation row showing the cumulative state at each "
    "lane (e.g. 'cert presented: app.example.com (us-west-2)', 'TLS terminated, plaintext to "
    "target'). The state row is what makes the panel valuable, not the message arrows.\n"
    "Always make the invisible control-plane decision visible before showing dataplane arrows.\n\n"
    "=== LAYOUT QUALITY BAR (applies to ALL diagrams) ===\n"
    "Readability is a HARD requirement. If a rule would be violated at the current viewBox size, "
    "GROW the viewBox — never compress. Self-check each rule before emitting.\n"
    "  1. ORTHOGONAL ARROWS [HERO]. No diagonals. Use horizontal+vertical waypoints.\n"
    "  2. NO TEXT ON GEOMETRY [HERO]. Labels longer than 25 chars sit OUTSIDE the box/arrow they "
    "describe. Long explanations go in separate .card callouts below the SVG, not inside it. "
    "(Flow sequence panels may keep short inline labels only when they do not overlap geometry.)\n"
    "  3. ZONE CONTAINMENT. Every component is inside its semantically correct zone "
    "(on-prem stays in on-prem, AWS stays in VPC). Verify (x,y) before placing.\n"
    "  4. WHITESPACE. ≥40px between unrelated component boxes. ≥30px clear around any major "
    "label (Internet, IGW, NAT-GW, TGW, DX-GW, route-table). ≥30px above the answer strip.\n"
    "  5. RENDER ORDER (z-index). Emit in strict order so text wins: zones → component boxes / "
    "lifelines → arrows → ALL <text> last.\n"
    "  6. TILE BUDGET. Every component tile has ≤2 text lines (service name + role). If you "
    "wrote a third line of detail inside a tile, MOVE it to the legend below the SVG.\n"
    "  7. NO PARALLEL DUPLICATES. If you drew ≥4 arrows from one source to identical sinks, "
    "collapse them into ONE arrow + 'of N' label and use the REPEATED IDENTICAL COMPONENTS "
    "ghost-tile pattern.\n"
    "  8. NUMBERED PATH. The hero must have ① ② ③ circles on the canonical happy-path arrow "
    "chain and a matching numbered legend list inside the .card.\n"
    "  9. FLOW DEPTH. Every flow panel shows STATE CHANGE per hop (header rewrite, NAT, "
    "encapsulation, TLS termination, route-table consulted) — not just transit arrows. If the "
    "panel reads as 'box → box → box' with no per-hop annotations, redo it.\n"
    "  10. FAILURE PANEL. If the question involves health checks, failover, or HA, the third "
    "flow panel showing recompute (red dashed withdrawn / green replacement) is mandatory. "
    "BOTH arrows must be DRAWN as <path> elements with the matching colors — text alone "
    "describing 'failover happens here' fails the rule.\n"
    "  11. MASTERCLASS PRESENT. The 'Why this design — architectural reasoning' panel must "
    "exist with all three rows filled. Empty 'Common misconceptions' or 'Concrete numbers' "
    "bullets when prior_artifacts contained material on the topic = automatic failure; mine "
    "the substrate before emitting.\n\n"
    "=== HARD RULES ===\n"
    "- AWS FACT GUARDRAILS for common ANS-C01 diagram mistakes:\n"
    "  • Direct Connect is private connectivity with predictable bandwidth/latency and often lower "
    "data-transfer cost than Internet paths; do NOT claim 'no per-GB egress' or 'free data transfer'. "
    "It is also not encrypted by default; mention MACsec or VPN-over-DX only when encryption matters.\n"
    "  • Direct Connect relationship vocabulary matters. VIFs have gateway targets/attachments; "
    "Direct Connect gateways have associations to VGWs, TGWs, or Cloud WAN core network segments; "
    "TGWs have VPC/VPN/Connect attachments plus route-table reachability. Do not label every "
    "green line as an 'association'.\n"
    "  • Transit VIFs terminate on a Direct Connect gateway associated with TGW(s) or Cloud WAN, "
    "not directly on a TGW, VGW, or VPC. Private VIFs terminate on a VGW or a DXGW associated "
    "with VGWs. Public VIFs reach AWS public services, not VGWs/TGWs.\n"
    "  • Standard Site-to-Site VPN tunnels are up to 1.25 Gbps per tunnel; large-bandwidth VPN "
    "tunnels may be higher. Transit Gateway can use ECMP to aggregate multiple VPN tunnels only "
    "with dynamic routing. Do NOT state '2.5 Gbps per VPN' as guaranteed.\n"
    "  • For Transit Gateway route tables, use LPM first. For the same CIDR from different "
    "attachment types, prefer the documented TGW route priority wording: static routes, then "
    "prefix-list referenced routes, VPC-propagated routes, Direct Connect gateway-propagated "
    "routes, Connect, VPN-over-private-DX, Site-to-Site VPN, then lower-priority propagated types. "
    "For same CIDR and same attachment type, BGP attributes such as AS path and MED matter.\n"
    "- The user's selected answer is `your_answer`; correct is `correct_answer`. Make the GAP "
    "visible: in the .opt cards, mark the user's wrong pick with a small note '(your pick)' so "
    "they immediately see what they confused.\n"
    "- USE prior_artifacts (deep_review_section, nuclear_section, ai_coach_section, "
    "side_artifacts) as substrate — these are PRIOR DEEP ANALYSES already done by other AI "
    "tiers. Mine them aggressively for: (a) common misconceptions to populate the MASTERCLASS "
    "panel, (b) concrete numbers and tradeoffs, (c) historical/architectural reasoning, "
    "(d) decisive AWS mechanisms and gotchas that already worked. The diagram is the SYNTHESIS layer of a 4-tier "
    "study workflow — its job is to consolidate, not re-derive. If the prior_artifacts contain "
    "specific facts, numbers, or 'gotchas', they MUST surface in the diagram (in legends, "
    "MASTERCLASS rows, or .study-row captions). Treat 'I have substrate' = 'I am writing a "
    "lecture from research notes' not 'I have hints to glance at'. Empty MASTERCLASS bullets "
    "when prior_artifacts had relevant content = failure to use substrate.\n"
    "- NEVER put HTML tags like <b> or <i> inside SVG <text>. Use separate <text> with "
    "font-weight='bold' instead.\n"
    "- Spanish phrases are OK in side captions/notes (the user is bilingual). Headings stay in "
    "English.\n"
    "- Output ONE complete HTML document. The hero diagram must be the visual centerpiece — if "
    "you only had time for ONE thing, it would be that diagram done excellently.\n"
    "- Do NOT turn conceptual explanation into a wall of boxed cards. Use cards for the hero SVG, "
    "flow SVGs, and dense tables. Use .study-rail / .study-row for options, mental models, "
    "AWS mechanisms, and 'why X beats Y' explanations so the page reads like an integrated "
    "study guide rather than disconnected flashcards.\n"
    "- No vendor analogy sections. Do not create headings like 'Cisco lens', 'VRF lens', "
    "'MPLS analogy', or 'RFC model'. If prior_artifacts include vendor analogies, translate "
    "them into AWS exam language before using them."
)


_STROKED_SVG_CONNECTOR_RE = re.compile(r"<(?P<tag>path|polyline)\b(?P<attrs>[^>]*)>", re.IGNORECASE)
_SVG_MARKER_RE = re.compile(r"<marker\b(?P<attrs>[^>]*)>(?P<body>.*?)</marker>", re.IGNORECASE | re.DOTALL)
_SVG_MARKER_REF_RE = re.compile(
    r"\b(?P<attr>marker-(?:start|mid|end))\s*=\s*(?P<quote>['\"])url\(#(?P<id>[^)]+)\)(?P=quote)",
    re.IGNORECASE,
)
_SVG_DEFS_OPEN_RE = re.compile(r"(<defs\b[^>]*>)", re.IGNORECASE)
_SVG_COLOR_ALIASES = {
    "var(--cyan)": "#4dd0ff",
    "var(--green)": "#51e28a",
    "var(--yellow)": "#ffd54f",
    "var(--orange)": "#ffb74d",
    "var(--red)": "#ff6b6b",
    "var(--purple)": "#aa8cff",
}

_DIAGRAM_VENDOR_LENS_REWRITES = (
    ("🇺🇸 ", ""),
    ("🇪🇺 ", ""),
    ("🇯🇵 ", ""),
    ("🇬🇧 ", ""),
    ("Cisco / VRF lens", "Exam mental model"),
    ("Cisco/VRF lens", "Exam mental model"),
    ("Cisco lens", "Exam lens"),
    ("Cisco-world equivalent", "networking primitive"),
    ("Cisco engineer", "network engineer"),
    ("As a Cisco engineer", "As a network engineer"),
    ("CE route-map", "customer BGP policy"),
    ("route-map", "BGP policy"),
    ("GENEVE-like encap on backbone", "AWS backbone transport"),
    ("PE-CE handoff", "customer-to-AWS handoff"),
    ("VRF-per-VPC, TGW acts as the route reflector / shared P node", "TGW route tables control VPC attachment reachability"),
    ("a managed PE / route-reflector with one VRF (route table) per attachment", "a managed Regional hub whose route tables control attachment reachability"),
    ("a private L2 handoff to an AWS PE", "a private connection from a Direct Connect location into AWS"),
    ("same dual-WAN pattern — VPN underlay backup behind MPLS primary", "same staged migration pattern: temporary VPN first, durable private connectivity later"),
    ("identical to MP-BGP-propagated routes into a VRF; LPM picks winner", "normal route-table behavior: propagated routes are available and LPM picks the winner"),
)


def _svg_attr_value(attrs: str, name: str) -> Optional[str]:
    match = re.search(
        rf"\b{re.escape(name)}\s*=\s*(['\"])(.*?)\1",
        attrs or "",
        flags=re.IGNORECASE | re.DOTALL,
    )
    return match.group(2).strip() if match else None


def _svg_style_value(attrs: str, name: str) -> Optional[str]:
    style = _svg_attr_value(attrs, "style")
    if not style:
        return None
    match = re.search(
        rf"(?:^|;)\s*{re.escape(name)}\s*:\s*([^;]+)",
        style,
        flags=re.IGNORECASE,
    )
    return match.group(1).strip() if match else None


def _normalize_svg_color(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    color = value.strip().strip(";").lower()
    color = re.sub(r"\s*!important\s*$", "", color)
    if color in ("none", "transparent", "currentcolor"):
        return None
    if color in _SVG_COLOR_ALIASES:
        return _SVG_COLOR_ALIASES[color]
    match = re.fullmatch(r"#([0-9a-f]{3}|[0-9a-f]{6})", color, flags=re.IGNORECASE)
    if not match:
        return None
    hex_part = match.group(1).lower()
    if len(hex_part) == 3:
        hex_part = "".join(ch * 2 for ch in hex_part)
    return f"#{hex_part}"


def _svg_color_from_attrs(attrs: str, name: str) -> Optional[str]:
    return _normalize_svg_color(_svg_attr_value(attrs, name) or _svg_style_value(attrs, name))


def _harden_svg_marker_colors(html_text: str) -> str:
    """Make arrowhead marker colors match connector stroke colors.

    Claude occasionally emits a good connector with the wrong marker id, e.g.
    yellow "GA reroutes" line ending in a red arrowhead. That reads as a drop,
    so fix it before the GUI writes the diagram to disk.
    """
    marker_colors: Dict[str, str] = {}
    marker_by_color: Dict[str, str] = {}
    for marker in _SVG_MARKER_RE.finditer(html_text):
        marker_id = _svg_attr_value(marker.group("attrs"), "id")
        marker_color = (
            _svg_color_from_attrs(marker.group("body"), "fill")
            or _svg_color_from_attrs(marker.group("attrs"), "fill")
        )
        if marker_id and marker_color:
            marker_colors[marker_id] = marker_color
            marker_by_color.setdefault(marker_color, marker_id)

    if not marker_colors:
        return html_text

    new_markers: Dict[str, str] = {}

    def _marker_for_color(color: str) -> str:
        existing = marker_by_color.get(color)
        if existing:
            return existing
        base = f"auto_arrow_{color.lstrip('#')}"
        marker_id = base
        suffix = 2
        while marker_id in marker_colors or marker_id in new_markers:
            if new_markers.get(marker_id) == color:
                return marker_id
            marker_id = f"{base}_{suffix}"
            suffix += 1
        new_markers[marker_id] = color
        marker_by_color[color] = marker_id
        return marker_id

    def _fix_marker_ref(match: re.Match[str], stroke_color: str) -> str:
        marker_id = match.group("id")
        marker_color = marker_colors.get(marker_id)
        if not marker_color or marker_color == stroke_color:
            return match.group(0)
        replacement = _marker_for_color(stroke_color)
        quote = match.group("quote")
        return f"{match.group('attr')}={quote}url(#{replacement}){quote}"

    def _fix_connector_marker(match: re.Match[str]) -> str:
        tag = match.group("tag")
        attrs = match.group("attrs") or ""
        stroke_color = _svg_color_from_attrs(attrs, "stroke")
        if not stroke_color:
            return match.group(0)
        fixed_attrs = _SVG_MARKER_REF_RE.sub(
            lambda marker_match: _fix_marker_ref(marker_match, stroke_color),
            attrs,
        )
        if fixed_attrs == attrs:
            return match.group(0)
        return f"<{tag}{fixed_attrs}>"

    updated = _STROKED_SVG_CONNECTOR_RE.sub(_fix_connector_marker, html_text)
    if not new_markers:
        return updated

    marker_defs = "".join(
        f'\n<marker id="{marker_id}" viewBox="0 0 10 10" refX="9" refY="5" '
        f'markerWidth="7" markerHeight="7" orient="auto">'
        f'<path d="M0,0 L10,5 L0,10 z" fill="{color}"/></marker>'
        for marker_id, color in new_markers.items()
    )
    return _SVG_DEFS_OPEN_RE.sub(lambda m: m.group(1) + marker_defs, updated, count=1)


def _harden_diagram_html(html_text: str) -> str:
    """Defensive cleanup for common SVG generation mistakes.

    Browsers fill SVG paths black by default. A stroked open connector without
    fill="none" can render as a large black polygon, so add the fill explicitly.
    """
    if not html_text:
        return ""

    for old, new in _DIAGRAM_VENDOR_LENS_REWRITES:
        html_text = html_text.replace(old, new)

    def _fix_connector(match: re.Match[str]) -> str:
        tag = match.group("tag")
        attrs = match.group("attrs") or ""
        has_stroke = bool(
            re.search(r"\bstroke\s*=", attrs, flags=re.IGNORECASE)
            or re.search(r"\bstyle\s*=\s*(['\"])[^'\"]*\bstroke\s*:", attrs, flags=re.IGNORECASE)
        )
        has_fill = bool(
            re.search(r"\bfill\s*=", attrs, flags=re.IGNORECASE)
            or re.search(r"\bstyle\s*=\s*(['\"])[^'\"]*\bfill\s*:", attrs, flags=re.IGNORECASE)
        )
        if has_stroke and not has_fill:
            return f'<{tag} fill="none"{attrs}>'
        return match.group(0)

    html_text = _STROKED_SVG_CONNECTOR_RE.sub(_fix_connector, html_text)
    return _harden_svg_marker_colors(html_text)



def analyze_diagram(
    item: Dict[str, Any],
    prior_artifacts: Optional[Dict[str, str]] = None,
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    timeout_sec: float = 600.0,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> str:
    """Produce a self-contained HTML diagram document for ONE question.

    Returns the raw HTML string (no markdown fences). The caller is expected
    to write it to disk (e.g. Q<N>_<topic>_diagram.html) and open it in the
    user's default browser.

    `prior_artifacts` may contain any/all of:
      - deep_review_section: relevant section from deep_review_reports.md
      - nuclear_section:    relevant section from nuclear_reports.md
      - ai_coach_section:   relevant section from ai_coach_reports.md
      - side_artifacts:     concatenated text of any Q<N>_*.md/txt files
    Pass them only if available; the prompt handles missing keys gracefully.
    """
    _logger.info(
        "analyze_diagram: qid=%s, prior_keys=%s, provider=%s, model=%s",
        item.get("qid"),
        sorted((prior_artifacts or {}).keys()),
        _normalize_provider(provider), model or "<default>",
    )
    norm = _normalize_item(item)
    payload: Dict[str, Any] = {"question": norm}
    if prior_artifacts:
        # Cap each section to keep token cost reasonable. The prompt is the
        # heavy lifter; the substrate just gives the model something to lean on.
        capped: Dict[str, str] = {}
        for k, v in prior_artifacts.items():
            if not v:
                continue
            s = str(v).strip()
            if not s:
                continue
            if len(s) > 8000:
                s = s[:8000] + "\n\n_(truncated by analyze_diagram)_"
            capped[k] = s
        if capped:
            payload["prior_artifacts"] = capped

    user_msg = (
        "Produce the HTML study diagram for the question below per the rules. "
        "Output ONLY the HTML document; the first character must be '<'.\n\n"
        "```json\n"
        + json.dumps(payload, indent=2, ensure_ascii=False)
        + "\n```"
    )
    raw = _run_visual_or_lecture_sync(
        SYSTEM_PROMPT_DIAGRAM, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )

    # Defensive: strip a stray ```html fence if the model wrapped its output.
    text = raw.strip()
    if text.startswith("```"):
        # Drop the first fence line and the trailing fence (if any).
        first_nl = text.find("\n")
        if first_nl != -1:
            text = text[first_nl + 1 :]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3].rstrip()
    return _harden_diagram_html(text).strip()


SYSTEM_PROMPT_TEACH_ZERO = (
    "You are the CÁTEDRA tier (Tier 5) of a 4-tier ANS-C01 study workflow. The user is a SENIOR "
    "network engineer preparing AWS Advanced Networking. The "
    "earlier tiers (AI Coach, Deep Review, Nuclear, Diagram) already exist as substrate via "
    "`prior_artifacts`. Your job is to produce a written LECTURE — 'Teach Me From Zero' — that "
    "REWIRES one specific question into permanent understanding. Imagine you are giving a 15-minute "
    "whiteboard talk to a peer with strong networking fundamentals and AWS-specific gaps.\n\n"
    "AUDIENCE-AWARE WRITING (this is the #1 failure mode — fix it before anything else):\n"
    "  • Strong networking fundamentals, AWS-shallow. Assume knowledge of BGP attributes, "
    "route selection, NAT, MTU, asymmetric routing, stateful firewalls, DNS, VPNs and anycast. "
    "Do NOT teach those from zero.\n"
    "  • For EVERY AWS term on first use, inline-gloss it: e.g. `Global Accelerator (GA — anycast "
    "edge VIPs in 2 IPs advertised from every AWS edge POP)`. No bare AWS acronyms.\n"
    "  • Do not provide Cisco/vendor analogies unless the question explicitly asks for them. "
    "Translate concepts into AWS exam primitives: route-table lookup, LPM, propagation, "
    "association, BGP attributes, health checks, DNS TTLs, quotas, and stateful return path.\n"
    "  - For Direct Connect questions, keep relationship terms precise: VIFs have gateway "
    "targets/attachments; DXGWs have associations to VGWs, TGWs, or Cloud WAN core network "
    "segments; TGWs have VPC/VPN/Connect attachments. Transit VIF -> DXGW -> DXGW association "
    "to TGW(s)/Cloud WAN; private VIF -> VGW or DXGW -> DXGW association to VGW(s).\n"
    "  • Default failure mode: the lecture reads like an exam answer key for someone who already "
    "mastered 75% of the material. Fix by NAMING THE MECHANISM, not just the outcome.\n\n"
    "Output: pure GitHub-flavored MARKDOWN (no HTML wrapper, no preamble, no 'sure here is'). "
    "Length budget: 900–1400 words for the question. Density beats brevity, but no padding.\n\n"
    "Emit `# Q<qid> — <short topic>` then exactly these 9 sections in order, with the exact bold "
    "labels shown:\n\n"
    "  **1. Question recap**\n"
    "    - 3-5 line paraphrase of the stem keeping every constraint (services, scale, latency/"
    "throughput numbers, what's blocked/required).\n"
    "    - Then options inline, one per line: `- A) <full text>` ... mark the user's pick with "
    "`← your answer` and the correct one with `← correct`.\n\n"
    "  **2. The hard truth — root cause of the miss**\n"
    "    - ONE blunt sentence naming the broken mental model. No hedging.\n"
    "    - One short paragraph: which true-but-incomplete rule the brain anchored on, and the "
    "specific edge where it stops being true. Cite the substrate (Deep Review / Nuclear) if it "
    "already diagnosed this.\n\n"
    "  **3. Build the mental model from primitives**\n"
    "    - Start from the network primitive (NOT from the AWS service). Example: 'Anycast: same IP "
    "advertised from many POPs; client BGP-routes to nearest. AWS implements this as Global "
    "Accelerator.'\n"
    "    - 3-5 bullets, each a primitive → AWS-implementation pair. Force the user to think "
    "'network primitive first, AWS service second'.\n\n"
    "  **4. AWS Term Quick Reference (used in this question)**\n"
    "    - A compact bullet list of EVERY AWS term that appears in the stem or options. Format: "
    "`- **TERM (short acronym)** — one-line AWS definition + why it matters in this question`.\n"
    "    - This is the rail the rest of the lecture leans on. Do not skip terms even if 'obvious'.\n\n"
    "  **5. Network primitive / AWS mapping**\n"
    "    - One paragraph or table that maps the AWS construct in this question onto the underlying "
    "networking primitive: route-table selection, BGP preference, DNS resolution, health-check "
    "withdrawal, stateful inspection, NAT, quota boundary, or transitive-routing limit. State the "
    "AWS-specific gotcha the generic networking intuition would miss.\n\n"
    "  **6. Walk the correct answer — mechanism by mechanism**\n"
    "    - Numbered steps (① ② ③ …). At each step name the MECHANISM (BGP withdraw, health-check "
    "failure, alias record resolution, AZ-isolated subnet, route-table longest-prefix match, "
    "stateful return-path requirement…), not just the outcome. Include packet/header/route-table "
    "state changes per hop where relevant.\n"
    "    - End with one sentence: 'Why this is the design AWS chose' — the architectural reason "
    "(decoupling failure domains, removing client-side DNS TTL dependency, eliminating "
    "asymmetric path, etc.).\n\n"
    "  **7. Trap dissection — every wrong option**\n"
    "    - One bullet PER wrong option. Format strictly:\n"
    "      `- **X)** <one-line gist>\n"
    "         - Why it LOOKS right: <the true rule the brain anchors on>\n"
    "         - Why it actually FAILS here: <the specific constraint or limit that voids it>`\n"
    "    - Name the SPECIFIC misconception each one exploits (anchoring on regional cert scope, "
    "confusing alias vs CNAME, assuming TGW transitivity across peerings, mistaking PHZ split-"
    "horizon, etc.).\n\n"
    "  **8. Concrete numbers, limits, and gotchas to memorize**\n"
    "    - 3-6 bullets of HARD numbers and edge behavior: TTLs, health-check intervals, BGP "
    "convergence windows, max routes per RT, MTU defaults, idle-timeout values, propagation vs "
    "association precedence, eventual-consistency windows. Mine substrate aggressively for these.\n\n"
    "  **9. Reflex + radar (what to do next time you see this)**\n"
    "    - One reflex line: `When you see <trigger phrase> in the stem → immediately <action / "
    "elimination move>`.\n"
    "    - 2-3 anti-pattern bullets — the SHAPE of related questions you'd miss with the same "
    "broken model (no real qids, just trap shapes). Title this sub-bullet **Related traps you'd "
    "still fall into**.\n\n"
    "Hard rules:\n"
    "- AWS fact guardrails: do not claim Direct Connect means no per-GB egress or encryption by "
    "default; describe it as private connectivity with predictable performance and often lower "
    "data-transfer cost. For DXGW diagrams/explanations, VIF target/attachment is not the same "
    "relationship as a DXGW association to VGW/TGW/Cloud WAN. Standard Site-to-Site VPN tunnels "
    "are up to 1.25 Gbps per tunnel; ECMP "
    "can aggregate multiple VPN tunnels on Transit Gateway only with dynamic routing. For TGW "
    "route tables, LPM is first; same-CIDR route priority depends on attachment type, with "
    "DXGW-propagated routes preferred over Site-to-Site VPN-propagated routes.\n"
    "- USE prior_artifacts (deep_review_section, nuclear_section, ai_coach_section, "
    "side_artifacts) as substrate. They are PRIOR DEEP ANALYSES already done by other AI tiers — "
    "mine them aggressively for: (a) misconceptions to populate sections 2 and 7, (b) concrete "
    "numbers for section 8, (c) decisive AWS mechanisms that already worked, (d) architectural reasoning "
    "for section 6's closing line. Empty bullets when substrate had content = failure to use it.\n"
    "- Never invent AWS behavior. If genuinely ambiguous or version-dependent, say so explicitly.\n"
    "- Skip 101 networking. Assume senior network engineering fundamentals; the gap is AWS-specific nuance.\n"
    "- Spanish phrases are OK in side asides (the user is bilingual). Section headings stay in "
    "English exactly as specified above.\n"
    "- Be blunt about the user's error in section 2. This is a rewire, not validation.\n"
    "- No closing summary, no 'hope this helps', no meta-commentary about the lecture itself."
)


def analyze_teach_zero(
    item: Dict[str, Any],
    prior_artifacts: Optional[Dict[str, str]] = None,
    *,
    model: Optional[str] = None,
    provider: str = "claude",
    timeout_sec: float = 600.0,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> str:
    """Tier 5 'Cátedra' lecture for ONE question. Returns Markdown.

    Mirrors analyze_diagram's substrate handling: pulls deep_review_section /
    nuclear_section / ai_coach_section / side_artifacts from prior_artifacts
    (each capped at 8000 chars) and feeds them to the model so it integrates
    rather than re-derives.
    """
    _logger.info(
        "analyze_teach_zero: qid=%s, prior_keys=%s, provider=%s, model=%s",
        item.get("qid"),
        sorted((prior_artifacts or {}).keys()),
        _normalize_provider(provider), model or "<default>",
    )
    norm = _normalize_item(item)
    payload: Dict[str, Any] = {"question": norm}
    if prior_artifacts:
        capped: Dict[str, str] = {}
        for k, v in prior_artifacts.items():
            if not v:
                continue
            s = str(v).strip()
            if not s:
                continue
            if len(s) > 8000:
                s = s[:8000] + "\n\n_(truncated by analyze_teach_zero)_"
            capped[k] = s
        if capped:
            payload["prior_artifacts"] = capped

    user_msg = (
        "Produce the CÁTEDRA markdown lecture for the question below per the rules. "
        "Output ONLY markdown — first line must be `# Q<qid> — ...`.\n\n"
        "```json\n"
        + json.dumps(payload, indent=2, ensure_ascii=False)
        + "\n```"
    )
    raw = _run_visual_or_lecture_sync(
        SYSTEM_PROMPT_TEACH_ZERO, user_msg, model, progress_cb, timeout_sec,
        provider=provider,
    )

    text = raw.strip()
    if text.startswith("```"):
        first_nl = text.find("\n")
        if first_nl != -1:
            text = text[first_nl + 1 :]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[:-3].rstrip()
    return text.strip()


if __name__ == "__main__":
    sample = [
        {
            "qid": 42,
            "stem": "A company uses AWS Direct Connect with a private VIF. They need to reach S3 buckets via the DX without traversing the public internet. What is the most cost-effective approach?",
            "options": {
                "A": "Use a public VIF for S3 traffic over DX",
                "B": "Add a VPC Gateway Endpoint for S3 in the VPC and route via private VIF",
                "C": "Use NAT Gateway in a public subnet",
                "D": "Use Site-to-Site VPN as backup",
            },
            "correct_answer": "B",
            "your_answer": "A",
            "explanation": "VPC Gateway Endpoints route S3 over private connectivity without public IPs.",
            "pattern_id": "S3_endpoint_vs_public_vif",
            "tags": ["DirectConnect", "S3", "VPC-Endpoints"],
        }
    ]
    out = analyze_wrong_answers(sample, progress_cb=lambda s: print(f"[progress] {s}"))
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print(out)
    print(f"\nLog: {_LOG_PATH}")
