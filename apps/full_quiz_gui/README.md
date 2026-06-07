# Full PySide6 Quiz GUI

This is the full desktop application version copied from the working local app.
The code is based on the `ans_c01_quiz_gui_v2_counterfix2.py` file from the
local ACTIVE workspace, with one added integration point: `RAG KB`.

It is intentionally included as code only. Real quiz banks and override files
must stay local under:

```text
data/private/
```

## Run

From the repo root:

```powershell
python .\scripts\run_full_quiz_gui.py
```

Or directly:

```powershell
python .\apps\full_quiz_gui\ans_c01_quiz_gui_v2_counterfix2.py
```

## What This App Demonstrates

- deterministic GUI/state orchestration
- DOCX parsing
- wrong-answer and profile persistence
- Claude/OpenAI provider selection
- AI Coach and Deep Review workflows
- concept dossier generation
- multi-step Nuclear Review workflow
- RAG KB retrieval against sanitized/public knowledge docs
- HTML diagram generation
- Markdown report persistence

## RAG KB

The `RAG KB` button demonstrates the enterprise pattern without exposing a real
question bank. Inside the full GUI it uses the currently loaded DOCX as the
primary retrieval corpus, excluding the current question so the assistant can
pull related questions and explanations instead of simply echoing the same
item.

It can also retrieve supplemental context from:

```text
data/private/knowledge_docs.md
```

when that private local file exists; otherwise it uses:

```text
data/mock/network_docs.md
```

The standalone scripts keep the dry-run safety gate unless `AI_LIVE=1` is
explicitly enabled. The full GUI button uses the same selected backend as AI
Coach / Deep Review and asks for confirmation before sending the current
question and retrieved context.

## GitHub Safety

Do not commit:

- real quiz DOCX files
- override CSVs from real banks
- generated AI reports based on real question content
- screenshots with protected/private data
- API keys

Use mock data in `data/mock/` for public demos.
