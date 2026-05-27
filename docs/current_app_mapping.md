# Current App Mapping

The original Python/PySide6 app already has the hard part: controlled GenAI
workflows around real user state.

## What Already Exists

The full original GUI is included in:

```text
apps/full_quiz_gui/
```

| Current app piece | What it means technically |
| --- | --- |
| AI Coach | First-pass diagnostic workflow |
| Deep Review | Per-item deeper analysis |
| Meta-Coach | Analysis across accumulated history |
| Concept Dossiers | Reusable topic knowledge generated from prior reports |
| Nuclear Review | Multi-step specialist fan-out plus synthesis |
| Diagram generation | LLM-created HTML study artifact |
| Teach Zero / Catedra | Long-form teaching workflow |
| Claude/OpenAI switch | Provider abstraction |
| Markdown/HTML persistence | Audit trail and reusable context |
| RAG KB | Formal retrieval layer: chunks, citations, grounded prompt and safe SDK call |

## What Was Missing For "Enterprise RAG"

The current app reused context from local Markdown reports, but it did not yet
have a formal retrieval layer with:

- source documents
- chunk IDs
- embeddings/vector search
- metadata filters
- citations
- retrieval evaluation
- access-control thinking

That is what this ready folder adds as a clean next step.

## How To Explain It

Use this:

> My original app already had applied GenAI workflows and agent-style analysis.
> This version separates the knowledge layer into a RAG architecture: ingestion,
> chunking, retrieval, metadata, citations and workflow prompts.
