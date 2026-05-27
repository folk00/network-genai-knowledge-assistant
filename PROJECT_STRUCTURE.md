# Project Structure

This repo is intentionally small. The goal is to make the architecture easy to
explain in an demos, not to hide the idea behind a giant framework.

```text
ai-network-knowledge-assistant-ready/
|
|-- README.md
|   Main project overview and summary.
|
|-- run.py
|   Default launcher. Starts the full original PySide6 GUI.
|
|-- start_gui.bat
|   Windows convenience launcher for the same GUI.
|
|-- PROJECT_STRUCTURE.md
|   This file. Quick map of the repo.
|
|-- data/
|   Synthetic input data only.
|
|-- apps/
|   Runnable apps, including the full PySide6 GUI copied from the working tool.
|   The full GUI is the main product demo; the smaller GUI is only a RAG demo.
|
|-- docs/
|   Short human-readable notes: architecture, RAG, roadmap, demo pitch.
|
|-- scripts/
|   Small demos that can run from the command line.
|
|-- src/network_ai_assistant/
|   Python code. Split into RAG components and workflow orchestration.
|
|-- tests/
|   Simple tests for the retrieval layer.
```

## Code Layout

```text
src/network_ai_assistant/rag/
|
|-- domain.py
|   Data classes: SourceDocument, KnowledgeChunk, RetrievalResult, RAGAnswer.
|
|-- ingest.py
|   Loads synthetic Markdown docs and chunks them.
|
|-- simple_retriever.py
|   Local deterministic retriever. No API key needed.
|
|-- prompting.py
|   Builds grounded prompts with citations.
|
|-- langchain_adapter.py
|   Optional LangChain vector-store example.
```

```text
src/network_ai_assistant/llm/
|
|-- safety.py
|   Dry-run gate, .env loader and best-effort redaction.
|
|-- clients.py
|   OpenAI / Anthropic SDK adapters plus safe DryRunClient.
```

```text
src/network_ai_assistant/workflows/
|
|-- network_diagnosis.py
|   Example workflow that turns an incident into a grounded diagnostic prompt.
```

```text
apps/full_quiz_gui/
|
|-- ans_c01_quiz_gui_v2_counterfix2.py
|   Full PySide6 GUI from the original local app, with the RAG KB hook added.
|
|-- quiz_ai_coach.py
|   Claude/OpenAI workflow module used by the full GUI.
```

## Why This Structure Works

- `rag/` owns knowledge retrieval.
- `workflows/` owns use-case logic.
- `scripts/` proves it runs.
- `docs/` explains the architecture.
- `data/mock/` keeps the demo safe and clean.
