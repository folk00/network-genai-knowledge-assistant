# network_ai_assistant

Python package for the ready scaffold.

Keep this package small:

- `rag/` retrieves and prepares context.
- `llm/` safely calls provider SDKs.
- `workflows/` decides what to do with that context.

The LLM provider layer can be added later by adapting the existing
`quiz_ai_coach.py` Claude/OpenAI gateway pattern.
