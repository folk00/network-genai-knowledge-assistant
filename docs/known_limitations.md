# Known Limitations

This project is intentionally honest about what is implemented today and what is
next.

## Current Limitations

- The public RAG demo uses a deterministic keyword retriever, not a production
  vector database.
- The full GUI can use private local question banks, but those are intentionally
  excluded from GitHub.
- Generated reports are local artifacts and are not part of the public repo.
- The RAG KB button is designed for study and explanation workflows, not
  autonomous infrastructure changes.
- The safety scan is a lightweight guardrail, not a full DLP system.
- There is no hosted web service yet.
- There is no production tenant model yet.

## Why These Choices Are Acceptable

The goal is to demonstrate architecture and working software without publishing
protected material. The project shows:

- workflow orchestration
- provider abstraction
- retrieval and citations
- report persistence
- local/private data boundaries
- a clear path toward enterprise hardening

## Production Next Steps

If this were expanded into an enterprise tool, the next steps would be:

- FastAPI backend
- persistent vector store
- embeddings and hybrid retrieval
- tenant-aware metadata filtering
- authentication and authorization
- source versioning
- structured outputs
- evaluation data set
- CI/CD checks
- observability and audit logs
- human approval before any network change

## Honest Scope

This is not a production autonomous network agent. It is a working applied
GenAI assistant that demonstrates the engineering patterns required to build
one safely: controlled workflows, retrieval, citations, provider abstraction,
report persistence and human validation.
