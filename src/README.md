# Source Code

The code is split by responsibility:

```text
network_ai_assistant/rag/
  ingestion, chunking, retrieval, prompt construction

network_ai_assistant/workflows/
  use-case orchestration, such as network diagnosis
```

This mirrors the original app:

```text
GUI/state/orchestration -> AI module -> provider -> report
```

But makes the knowledge layer cleaner and easier to explain.

