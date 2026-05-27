from __future__ import annotations

import asyncio
import os
import sys
from dataclasses import dataclass
from typing import Protocol

from .safety import live_calls_enabled, redact_sensitive_text


class LLMClient(Protocol):
    def generate(self, prompt: str) -> str:
        """Return model output for a prompt."""


@dataclass(frozen=True)
class DryRunClient:
    """Safe default client. Never sends prompts to an external API."""

    provider: str = "dry-run"

    def generate(self, prompt: str) -> str:
        safe_prompt = redact_sensitive_text(prompt)
        preview = safe_prompt[:2500]
        if len(safe_prompt) > len(preview):
            preview += "\n\n...[prompt truncated in dry-run preview]..."
        return (
            "DRY RUN: no external LLM call was made.\n\n"
            "Set AI_LIVE=1 plus the provider API key to run a live SDK call.\n\n"
            "Prompt preview:\n"
            "----------------\n"
            f"{preview}"
        )


@dataclass(frozen=True)
class OpenAIClient:
    model: str

    def generate(self, prompt: str) -> str:
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise RuntimeError("Install OpenAI SDK: pip install openai") from exc

        if not os.environ.get("OPENAI_API_KEY"):
            raise RuntimeError("OPENAI_API_KEY is not set.")

        client = OpenAI(base_url=os.environ.get("OPENAI_BASE_URL") or None)
        response = client.responses.create(
            model=self.model,
            input=redact_sensitive_text(prompt),
            store=False,
        )
        text = getattr(response, "output_text", "")
        if not text:
            text = str(response)
        return text


@dataclass(frozen=True)
class ClaudeCodeClient:
    """Claude Code SDK client.

    This uses the local Claude Code authentication/session. No API key is read or
    stored by this project.
    """

    model: str | None = None
    timeout_sec: float = 180.0

    def generate(self, prompt: str) -> str:
        try:
            from claude_agent_sdk import AssistantMessage, ClaudeAgentOptions, TextBlock, query
        except ImportError as exc:
            raise RuntimeError("Install Claude Agent SDK: pip install claude-agent-sdk") from exc

        async def _run() -> str:
            options = ClaudeAgentOptions(
                system_prompt=(
                    "You are a network infrastructure assistant. Answer only from "
                    "the provided context, cite sources, and say when context is missing."
                ),
                allowed_tools=[],
                permission_mode="default",
                max_turns=1,
                setting_sources=[],
                mcp_servers={},
            )
            if self.model:
                options.model = self.model

            chunks: list[str] = []
            async for msg in query(prompt=redact_sensitive_text(prompt), options=options):
                if isinstance(msg, AssistantMessage):
                    for block in msg.content:
                        if isinstance(block, TextBlock):
                            chunks.append(block.text)
            return "".join(chunks).strip()

        if sys.platform == "win32":
            try:
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            except Exception:
                pass

        loop = asyncio.new_event_loop()
        try:
            asyncio.set_event_loop(loop)
            return loop.run_until_complete(asyncio.wait_for(_run(), timeout=self.timeout_sec))
        finally:
            try:
                pending = asyncio.all_tasks(loop)
                for task in pending:
                    task.cancel()
                if pending:
                    loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
            finally:
                loop.close()


@dataclass(frozen=True)
class AnthropicClient:
    model: str

    def generate(self, prompt: str) -> str:
        try:
            import anthropic
        except ImportError as exc:
            raise RuntimeError("Install Anthropic SDK: pip install anthropic") from exc

        if not os.environ.get("ANTHROPIC_API_KEY"):
            raise RuntimeError("ANTHROPIC_API_KEY is not set.")

        client = anthropic.Anthropic()
        response = client.messages.create(
            model=self.model,
            max_tokens=1600,
            messages=[{"role": "user", "content": redact_sensitive_text(prompt)}],
        )
        chunks: list[str] = []
        for block in response.content:
            text = getattr(block, "text", None)
            if text:
                chunks.append(text)
        return "\n".join(chunks).strip()


def build_client(provider: str | None = None) -> LLMClient:
    """Build an SDK client only when explicitly enabled.

    By default this returns DryRunClient to avoid accidental data exposure.
    """

    provider_name = (provider or os.environ.get("AI_PROVIDER") or "openai").strip().lower()
    if not live_calls_enabled():
        return DryRunClient(provider=provider_name)

    if provider_name in {"claude_code", "claude-code", "claudecode"}:
        model = (os.environ.get("CLAUDE_CODE_MODEL") or "").strip() or None
        timeout = float(os.environ.get("CLAUDE_CODE_TIMEOUT_SEC") or "180")
        return ClaudeCodeClient(model=model, timeout_sec=timeout)
    if provider_name == "openai":
        model = (os.environ.get("OPENAI_MODEL") or "gpt-4.1-mini").strip()
        return OpenAIClient(model=model)
    if provider_name in {"anthropic", "claude"}:
        model = (os.environ.get("ANTHROPIC_MODEL") or "").strip()
        if not model:
            raise RuntimeError("ANTHROPIC_MODEL must be set for live Anthropic calls.")
        return AnthropicClient(model=model)
    raise ValueError(f"Unsupported AI_PROVIDER: {provider_name}")
