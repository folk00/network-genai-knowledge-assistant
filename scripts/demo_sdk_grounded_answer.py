from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from network_ai_assistant.llm.clients import build_client
from network_ai_assistant.llm.safety import load_env_file
from network_ai_assistant.rag.ingest import chunk_documents, load_markdown_documents
from network_ai_assistant.rag.prompting import build_grounded_prompt
from network_ai_assistant.rag.simple_retriever import KeywordRetriever


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    load_env_file(ROOT / ".env")

    docs = load_markdown_documents(ROOT / "data" / "mock" / "network_docs.md")
    chunks = chunk_documents(docs)
    retriever = KeywordRetriever(chunks)

    question = (
        "A Cisco branch advertises 172.16.20.0/24 to AWS through a Site-to-Site "
        "VPN attached to Transit Gateway. What should I validate if the VPC cannot "
        "reach the branch LAN?"
    )
    results = retriever.search(question, k=3, filters={"tenant": "demo-retail"})
    prompt = build_grounded_prompt(question, results)

    client = build_client()
    answer = client.generate(prompt)
    print(answer)


if __name__ == "__main__":
    main()
