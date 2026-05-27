from __future__ import annotations

from pathlib import Path


def build_langchain_vectorstore(markdown_path: Path):
    """Example LangChain adapter for the same RAG pattern.

    This function is intentionally optional. The rest of the project can run
    without LangChain installed.

    Production upgrades:
    - replace InMemoryVectorStore with Chroma, OpenSearch, Pinecone, Azure AI
      Search, pgvector, or an enterprise-approved vector store
    - add metadata filters for tenant/site/security labels
    - add evaluation and tracing
    """

    try:
        from langchain_core.documents import Document
        from langchain_core.vectorstores import InMemoryVectorStore
        from langchain_openai import OpenAIEmbeddings
        from langchain_text_splitters import RecursiveCharacterTextSplitter
    except ImportError as exc:
        raise RuntimeError(
            "Install optional RAG dependencies first: "
            "pip install -r requirements.txt"
        ) from exc

    raw = markdown_path.read_text(encoding="utf-8")
    docs = [Document(page_content=raw, metadata={"source": str(markdown_path)})]

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=900,
        chunk_overlap=120,
    )
    chunks = splitter.split_documents(docs)

    embeddings = OpenAIEmbeddings(model="text-embedding-3-large")
    vector_store = InMemoryVectorStore(embedding=embeddings)
    vector_store.add_documents(chunks)
    return vector_store


def retrieve_with_langchain(vector_store, query: str, k: int = 4):
    """Return LangChain Documents from the vector store."""

    return vector_store.similarity_search(query, k=k)

