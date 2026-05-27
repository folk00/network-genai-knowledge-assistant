from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from network_ai_assistant.rag.ingest import chunk_documents, load_markdown_documents
from network_ai_assistant.rag.simple_retriever import KeywordRetriever


def test_retriever_finds_tgw_vpn_doc():
    docs = load_markdown_documents(ROOT / "data" / "mock" / "network_docs.md")
    chunks = chunk_documents(docs)
    retriever = KeywordRetriever(chunks)

    results = retriever.search("transit gateway vpn branch route 172.16.20.0/24", k=2)

    assert results
    assert results[0].chunk.source_id == "DOC-001"

