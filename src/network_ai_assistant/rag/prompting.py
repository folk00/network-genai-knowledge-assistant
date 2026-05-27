from __future__ import annotations

from .domain import RetrievalResult


def build_grounded_prompt(question: str, results: list[RetrievalResult]) -> str:
    """Build a prompt that forces source-grounded answers."""

    context_blocks = []
    for i, result in enumerate(results, start=1):
        chunk = result.chunk
        title = chunk.metadata.get("title", chunk.source_id)
        context_blocks.append(
            f"[{i}] source={chunk.source_id} chunk={chunk.chunk_id} title={title}\n"
            f"{chunk.text}"
        )

    context = "\n\n---\n\n".join(context_blocks) or "(no context retrieved)"
    return f"""You are a network infrastructure assistant.

Answer only from the retrieved context. If the context is insufficient, say what
is missing. Cite sources using [1], [2], etc.

Question:
{question}

Retrieved context:
{context}

Answer:
"""


def deterministic_summary(question: str, results: list[RetrievalResult]) -> str:
    """No-LLM fallback useful for demos and tests."""

    if not results:
        return "No relevant context was retrieved."
    lines = [f"Question: {question}", "", "Retrieved evidence:"]
    for i, result in enumerate(results, start=1):
        chunk = result.chunk
        title = chunk.metadata.get("title", chunk.source_id)
        lines.append(f"- [{i}] {chunk.source_id} / {title} / score={result.score:.3f}")
    lines.append("")
    lines.append("Next step: pass this prompt to Claude/OpenAI for a grounded answer with citations.")
    return "\n".join(lines)

