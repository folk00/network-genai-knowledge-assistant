from __future__ import annotations

import math
import re
from collections import Counter
from typing import Iterable

from .domain import KnowledgeChunk, RetrievalResult


TOKEN_RE = re.compile(r"[a-zA-Z0-9_.:/-]+")


class KeywordRetriever:
    """Small deterministic retriever for local demos.

    This is not a replacement for embeddings. It exists so the walkthrough demo can
    run without API keys or vector DB setup.
    """

    def __init__(self, chunks: Iterable[KnowledgeChunk]):
        self.chunks = list(chunks)
        self._chunk_terms = [Counter(_tokens(c.text)) for c in self.chunks]

    def search(
        self,
        query: str,
        *,
        k: int = 4,
        filters: dict[str, str] | None = None,
    ) -> list[RetrievalResult]:
        query_terms = Counter(_tokens(query))
        if not query_terms:
            return []

        results: list[RetrievalResult] = []
        for chunk, terms in zip(self.chunks, self._chunk_terms):
            if filters and not _metadata_matches(chunk.metadata, filters):
                continue
            score = _cosine(query_terms, terms)
            if score > 0:
                results.append(RetrievalResult(chunk=chunk, score=score))
        return sorted(results, key=lambda r: r.score, reverse=True)[:k]


def _tokens(text: str) -> list[str]:
    return [t.lower() for t in TOKEN_RE.findall(text or "") if len(t) > 2]


def _cosine(a: Counter[str], b: Counter[str]) -> float:
    shared = set(a) & set(b)
    numerator = sum(a[t] * b[t] for t in shared)
    if numerator == 0:
        return 0.0
    norm_a = math.sqrt(sum(v * v for v in a.values()))
    norm_b = math.sqrt(sum(v * v for v in b.values()))
    return numerator / (norm_a * norm_b)


def _metadata_matches(metadata: dict, filters: dict[str, str]) -> bool:
    for key, expected in filters.items():
        actual = str(metadata.get(key, "")).lower()
        if str(expected).lower() not in actual:
            return False
    return True

