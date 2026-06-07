# RAG And LangChain Notes

## RAG In Plain English

RAG means the app does not ask the model to answer from memory only.

Instead:

```text
user question
  -> retrieve relevant docs
  -> inject docs into the prompt
  -> ask the model to answer with citations
```

For network automation, the retrieved docs could be:

- configs
- MOPs
- incident notes
- runbooks
- topology notes
- previous AI reports
- diagrams
- AWS/TGW/VPN design notes

## Enterprise RAG

Enterprise RAG adds control:

- chunking strategy
- metadata filters
- tenant/site/source labels
- access controls
- citations
- evaluation
- tracing
- human validation

Key principle:

> RAG is not just a vector database. The real value is controlled retrieval,
> source grounding and auditability.

## Where LangChain Fits

LangChain gives reusable interfaces for:

- document loaders
- text splitters
- embedding models
- vector stores
- retrievers
- tools
- agent workflows

In this repo:

```text
simple_retriever.py      local no-dependency demo
langchain_adapter.py     optional LangChain version
```

That keeps two paths open:

- understand the architecture without depending on a framework
- plug it into LangChain-style components when needed

## Practical Implementation Path

1. Start with local Markdown/TXT ingestion.
2. Add chunking and metadata.
3. Use deterministic retrieval for demo.
4. Add LangChain + OpenAI embeddings.
5. Replace in-memory vector store with Chroma/OpenSearch/Pinecone/pgvector.
6. Add metadata filters and citations.
7. Add tests that prove the correct document is retrieved.

## Official References

- LangChain Retrieval: https://docs.langchain.com/oss/python/langchain/retrieval
- LangChain Vector Stores: https://docs.langchain.com/oss/python/integrations/vectorstores/
- LangChain Python Reference: https://reference.langchain.com/python/langchain/overview

