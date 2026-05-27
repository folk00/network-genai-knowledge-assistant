# Architecture

## Target Architecture

```text
                         +----------------------+
                         |  Network AI UI/API   |
                         |  PySide / FastAPI    |
                         +----------+-----------+
                                    |
                                    v
                         +----------+-----------+
                         | Workflow Orchestrator|
                         | deterministic logic  |
                         +----------+-----------+
                                    |
             +----------------------+----------------------+
             |                                             |
             v                                             v
 +-----------+------------+                    +-----------+------------+
 | Retrieval Service      |                    | LLM Provider Gateway   |
 | metadata filters       |                    | Claude / OpenAI / etc. |
 | citations              |                    | retry / timeout / logs |
 +-----------+------------+                    +-----------+------------+
             |                                             |
             v                                             v
 +-----------+------------+                    +-----------+------------+
 | Knowledge Index        |                    | Controlled Prompts     |
 | vector + keyword       |                    | structured output      |
 | chunks + metadata      |                    | human validation       |
 +-----------+------------+                    +-----------+------------+
             |
             v
 +-----------+------------+
 | Source Documents       |
 | MOPs, configs, logs,   |
 | diagrams, runbooks     |
 +------------------------+
```

## Core Design Principle

The LLM should not own the application.

The application owns:

- data ingestion
- tenant/source metadata
- retrieval filters
- prompt shape
- state
- audit logs
- report persistence
- user confirmation

The LLM owns:

- explanation
- synthesis
- diagnosis
- diagram wording
- recommendation drafts

## RAG Flow

1. Ingest source documents.
2. Normalize text and remove sensitive data.
3. Split into chunks.
4. Attach metadata:
   - `tenant_id`
   - `source_type`
   - `site`
   - `technology`
   - `security_level`
   - `created_at`
   - `source_uri`
5. Store chunks in retrieval index.
6. User asks a question or launches a workflow.
7. Retriever fetches top chunks using query + metadata filters.
8. Prompt builder injects retrieved context with citations.
9. LLM generates answer/report.
10. System persists output with source citations and validation status.

## Enterprise Controls

| Control | Why it matters |
| --- | --- |
| Metadata filters | Avoid retrieving data from the wrong tenant/site/source |
| Source citations | Make answers auditable |
| Chunk IDs | Trace every claim back to a document section |
| Prompt templates | Keep workflows predictable |
| Evaluation set | Test retrieval quality before demos/demos |
| Human approval | Avoid automatic changes from unverified LLM output |
| Secret redaction | Prevent leaks in prompts, logs and reports |

## Good scope distinction

Say:

> I have implemented RAG-style context reuse and I am extending it into a more
> enterprise-grade RAG pattern with indexing, metadata filtering, citations and
> evaluation.

Avoid:

> I am a production RAG platform expert.

