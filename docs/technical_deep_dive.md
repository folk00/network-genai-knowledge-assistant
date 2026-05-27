# Technical Deep Dive

This document explains how the project is structured and what technical choices
it demonstrates.

## Core Idea

The application is intentionally not model-first. It is workflow-first.

The deterministic application owns:

- input parsing
- user state
- selected scope
- retrieval
- prompt construction
- report persistence
- provider selection
- user confirmation

The LLM owns:

- explanation
- synthesis
- study coaching
- diagnosis-style reasoning
- diagram or lecture generation

This is the same split needed in infrastructure automation: software must own
state and guardrails; models can help reason over selected context.

## Full GUI

Location:

```text
apps/full_quiz_gui/
```

Important files:

```text
ans_c01_quiz_gui_v2_counterfix2.py
quiz_ai_coach.py
```

The GUI demonstrates:

- PySide6 application state
- DOCX parsing
- override CSV support
- profile persistence
- wrong-answer bank
- confidence tracking
- Markdown/HTML report artifacts
- Claude/OpenAI provider switching
- multi-step analysis workflows

## RAG Layer

Location:

```text
src/network_ai_assistant/rag/
```

Components:

| File | Purpose |
| --- | --- |
| `domain.py` | Data classes for sources, chunks and retrieval results |
| `ingest.py` | Loads Markdown source docs and splits them into chunks |
| `simple_retriever.py` | Deterministic keyword retriever for local demos |
| `prompting.py` | Builds grounded prompts with citations |
| `langchain_adapter.py` | Optional LangChain/vector-store implementation path |

The simple retriever is not presented as production search. It exists so the
demo can run without API keys, embeddings or vector DB setup.

## RAG KB In The GUI

The full GUI has a `RAG KB / Grounded Context` button.

Runtime behavior:

1. Read the current question from GUI state.
2. Build temporary source documents from the loaded question bank.
3. Exclude the current question from retrieval.
4. Search for related question-bank context.
5. Optionally add supplemental context from `data/private/knowledge_docs.md` or
   `data/mock/network_docs.md`.
6. Build a grounded prompt with citations.
7. Ask for confirmation.
8. Send one call to the selected backend using the same mechanism as the older
   AI buttons.
9. Render a study review dialog and keep the full grounded prompt in details.

This demonstrates how a personal app can evolve toward enterprise RAG without
turning the whole application into an unstructured chat interface.

## LLM Provider Layer

Location:

```text
src/network_ai_assistant/llm/
```

The standalone demos use a safe client builder:

- dry-run by default
- live only when `AI_LIVE=1`
- no credentials stored in code
- best-effort prompt redaction

The full GUI uses the existing backend path in `quiz_ai_coach.py` so the user
experience matches AI Coach, Deep Review, Diagram and Teach Zero.

## Report Persistence

The original app writes Markdown and HTML artifacts such as:

- AI coach reports
- deep review reports
- nuclear review reports
- concept dossiers
- diagrams
- teach-zero lectures

Those generated artifacts are intentionally ignored in Git because they can
contain private question text or private analysis.

## Why This Maps To Network Operations

For study:

```text
question + wrong answer + concept
  -> retrieved context
  -> LLM review
  -> persisted study artifact
```

For operations:

```text
incident + logs + config + runbook
  -> retrieved evidence
  -> LLM diagnosis
  -> validation checklist / rollback notes
  -> persisted operations artifact
```

The workflow shape is the same. The source data changes.

## Production Hardening Ideas

If this became a larger production-style project, next steps would be:

- FastAPI service boundary
- persistent vector store
- embeddings with hybrid keyword/vector retrieval
- per-tenant access filters
- source versioning
- retrieval evaluation set
- structured LLM outputs
- audit trail with source citations
- CI checks for safety scan and tests
- sanitized demo data generator
