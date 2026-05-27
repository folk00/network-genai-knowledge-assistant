# Safe SDK Usage

This repo can call Claude Code, OpenAI or Anthropic, but it is safe by default.

## Default Behavior

By default, scripts run in dry-run mode:

```powershell
python .\scripts\demo_sdk_grounded_answer.py
```

Dry-run mode does not send anything to an external API. It prints the grounded
prompt preview so you can inspect what would be sent.

## Live OpenAI Call

Use only mock/sanitized data.

```powershell
$env:AI_LIVE="1"
$env:AI_PROVIDER="openai"
$env:OPENAI_API_KEY="sk-..."
$env:OPENAI_MODEL="gpt-4.1-mini"
python .\scripts\demo_sdk_grounded_answer.py
```

## Live Claude Code SDK Call

This uses your local Claude Code authentication. No API key is stored in this
project.

```powershell
$env:AI_LIVE="1"
$env:AI_PROVIDER="claude_code"
python .\scripts\demo_sdk_grounded_answer.py
```

Optional:

```powershell
$env:CLAUDE_CODE_MODEL="<your-claude-code-model-name>"
$env:CLAUDE_CODE_TIMEOUT_SEC="240"
```

## Live Anthropic Call

Use only mock/sanitized data.

```powershell
$env:AI_LIVE="1"
$env:AI_PROVIDER="anthropic"
$env:ANTHROPIC_API_KEY="sk-ant-..."
$env:ANTHROPIC_MODEL="<your-anthropic-model-name>"
python .\scripts\demo_sdk_grounded_answer.py
```

## Safety Rules

- Keep `AI_LIVE=0` unless you intentionally want a real API call.
- Never put real API keys in Git.
- Claude Code auth stays outside this project.
- Never send real certification question banks.
- Never send customer configs or production logs without sanitization.
- Keep `.env` local. The repo includes `.env.example` only.
- The SDK layer performs best-effort redaction before API calls, but redaction is
  not a replacement for using clean inputs.
