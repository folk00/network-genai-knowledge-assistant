from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class SourceDocument:
    """Original source material before chunking."""

    source_id: str
    text: str
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class KnowledgeChunk:
    """Retrievable chunk with metadata and source traceability."""

    chunk_id: str
    source_id: str
    text: str
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class RetrievalResult:
    """Chunk returned by retrieval with a deterministic score."""

    chunk: KnowledgeChunk
    score: float


@dataclass(frozen=True)
class RAGAnswer:
    """Grounded answer plus citations."""

    answer: str
    citations: list[RetrievalResult]

