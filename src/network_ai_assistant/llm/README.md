# LLM Layer

This folder owns safe SDK access.

Design rules:

- dry-run by default
- API keys only from environment variables
- no keys stored in code or settings
- prompts are redacted before logging
- live calls require `AI_LIVE=1`
- demo data only

Supported providers:

- Claude Code SDK (`claude_agent_sdk`)
- OpenAI SDK
- Anthropic SDK

The rest of the app should call this layer instead of talking directly to an SDK.
