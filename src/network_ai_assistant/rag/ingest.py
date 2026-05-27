from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

from .domain import KnowledgeChunk, SourceDocument


DOC_HEADER_RE = re.compile(r"^##\s+(DOC-\d+)\s+-\s+(.+)$", re.MULTILINE)


def load_markdown_documents(path: Path) -> list[SourceDocument]:
    """Load synthetic Markdown docs separated by DOC headers.

    This loader is intentionally simple for walkthrough demos. A production loader
    would handle DOCX/PDF/HTML, redaction, source versioning and access labels.
    """

    text = path.read_text(encoding="utf-8")
    matches = list(DOC_HEADER_RE.finditer(text))
    docs: list[SourceDocument] = []
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        source_id = match.group(1)
        title = match.group(2).strip()
        body = text[start:end].strip()
        metadata = _extract_metadata(body)
        metadata["title"] = title
        docs.append(SourceDocument(source_id=source_id, text=body, metadata=metadata))
    return docs


def chunk_documents(docs: Iterable[SourceDocument], max_chars: int = 900) -> list[KnowledgeChunk]:
    """Split docs into small chunks while preserving source metadata."""

    chunks: list[KnowledgeChunk] = []
    for doc in docs:
        paragraphs = [p.strip() for p in re.split(r"\n\s*\n", doc.text) if p.strip()]
        current: list[str] = []
        current_len = 0
        chunk_no = 1
        for paragraph in paragraphs:
            if current and current_len + len(paragraph) > max_chars:
                chunks.append(_make_chunk(doc, chunk_no, current))
                chunk_no += 1
                current = []
                current_len = 0
            current.append(paragraph)
            current_len += len(paragraph)
        if current:
            chunks.append(_make_chunk(doc, chunk_no, current))
    return chunks


def _make_chunk(doc: SourceDocument, chunk_no: int, parts: list[str]) -> KnowledgeChunk:
    metadata = dict(doc.metadata)
    metadata["chunk_no"] = chunk_no
    return KnowledgeChunk(
        chunk_id=f"{doc.source_id}:{chunk_no}",
        source_id=doc.source_id,
        text="\n\n".join(parts),
        metadata=metadata,
    )


def _extract_metadata(text: str) -> dict[str, str]:
    metadata: dict[str, str] = {}
    for key in ("Tenant", "Site", "Technology"):
        match = re.search(rf"^{key}:\s*(.+)$", text, flags=re.MULTILINE)
        if match:
            metadata[key.lower()] = match.group(1).strip()
    return metadata

