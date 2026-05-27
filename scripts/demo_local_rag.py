from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from network_ai_assistant.rag.ingest import chunk_documents, load_markdown_documents
from network_ai_assistant.rag.prompting import deterministic_summary, build_grounded_prompt
from network_ai_assistant.rag.simple_retriever import KeywordRetriever


def main() -> None:
    docs_path = ROOT / "data" / "mock" / "network_docs.md"
    docs = load_markdown_documents(docs_path)
    chunks = chunk_documents(docs)
    retriever = KeywordRetriever(chunks)

    question = (
        "For a branch advertising 172.16.20.0/24 over AWS Site-to-Site VPN "
        "to Transit Gateway, what routes must exist?"
    )
    results = retriever.search(
        question,
        k=3,
        filters={"tenant": "demo-retail"},
    )

    print("=== Deterministic retrieval summary ===")
    print(deterministic_summary(question, results))
    print()
    print("=== Prompt to send to Claude/OpenAI ===")
    print(build_grounded_prompt(question, results))


if __name__ == "__main__":
    main()

