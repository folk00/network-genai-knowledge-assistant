# Demo Walkthrough

This is the recommended demo path for a viewers screen or technical
conversation.

## 30-Second Intro

Say:

> This is a GenAI-assisted network knowledge assistant. The GUI is a real
> Python/PySide6 application that controls parsing, state, profiles, retrieval
> and reports. Claude/OpenAI-style models are used only inside controlled
> workflows such as AI Coach, Deep Review, RAG KB and diagram generation.

Then open:

```powershell
python run.py
```

## Demo Path

### 1. Show Deterministic App Logic

Open the full GUI and point out:

- DOCX loading
- override CSV support
- saved profiles
- wrong-answer tracking
- repeat-all questions
- confidence tracking
- per-question statistics

key point:

> The app is not just a prompt. It has deterministic learning state and workflow
> control around the model.

### 2. Show AI Coach / Deep Review

Use a safe local/private test file if you are demoing live. Do not show protected
question banks on a public screen recording.

key point:

> AI Coach and Deep Review are workflow buttons. The app builds the structured
> context, sends it to the selected backend, and persists the response as a
> Markdown artifact.

### 3. Show RAG KB

Click `RAG KB / Grounded Context`.

What happens:

- the current question is used as the query
- the current question is excluded from retrieval
- related items from the loaded bank are indexed temporarily
- supplemental safe docs can be included
- the model gets a grounded prompt with citations

key point:

> This demonstrates the RAG pattern I would use for network operations. Replace
> questions with MOPs, configs, logs and runbooks, and the workflow shape is the
> same.

### 4. Show The Public RAG Script

If you do not want to use private files:

```powershell
python .\scripts\demo_local_rag.py
```

key point:

> The public demo uses synthetic data and does not require an API key. It shows
> ingestion, chunking, retrieval and grounded prompt construction.

### 5. Show Safety Boundaries

Run:

```powershell
python .\scripts\safety_scan.py
```

key point:

> I separated private material from publishable code. The repo includes mock
> data, safety checks and Git ignore rules for private banks, reports, logs and
> credentials.

## What Not To Demo

Do not show:

- real certification question banks
- screenshots with protected question text
- customer configs
- API keys
- local auth/session files
- generated reports based on private content

## Best Closing Line

> The important part is the architecture: deterministic software controls data,
> context and workflow; the LLM provides reasoning inside clear boundaries.
