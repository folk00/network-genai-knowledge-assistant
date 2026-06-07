# Full GUI Integration Notes

This repo treats the original desktop app as the main product demo.

## Source Of Truth

The main GUI code comes from the local working file:

```text
C:\Users\davidgo2\Downloads\00_ACTIVE\ans_c01_quiz_gui_v2_counterfix2.py
```

The AI workflow module comes from:

```text
C:\Users\davidgo2\Downloads\00_ACTIVE\quiz_ai_coach.py
```

In this repo they live under:

```text
apps/full_quiz_gui/
```

## What Was Added

The GitHub-ready version keeps the original GUI behavior and adds one modular
integration point:

```text
RAG KB
```

That button connects the quiz GUI to the sanitized retrieval and SDK layer under:

```text
src/network_ai_assistant/
```

The old app remains useful as the polished UI. The new project structure adds
safe data boundaries, mock data, retrieval, citations, provider abstraction
and dry-run defaults.

## What Should Stay Private

Do not publish real exam banks, override files, generated reports, concept
dossiers from protected content, screenshots with real questions, API keys or
local state files.

Use these local-only paths for private material:

```text
data/private/
apps/full_quiz_gui/concept_dossiers/
apps/full_quiz_gui/*reports*.md
```
