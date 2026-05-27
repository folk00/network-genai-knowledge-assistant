# RAG Layer

This folder owns retrieval.

Current files:

- `domain.py` - data classes
- `ingest.py` - Markdown loader and chunker
- `simple_retriever.py` - local deterministic retriever
- `prompting.py` - grounded prompt builder
- `langchain_adapter.py` - optional LangChain vector-store example

Production next steps:

- add DOCX/PDF loaders
- add PII/secret redaction
- add embeddings
- add vector store
- add reranking
- add metadata access filters
- add retrieval evals

