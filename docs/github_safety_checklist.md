# GitHub Safety Checklist

Use this before publishing the repo.

## Safe To Upload

- source code
- architecture docs
- mock data
- sanitized screenshots
- synthetic generated reports
- README files

## Do Not Upload

- real certification question banks
- real answer keys
- DOCX quiz files
- override files built from real question banks
- customer configs
- production logs
- API keys
- `.env`
- generated reports based on real questions
- screenshots showing protected/private data

## Local Workflow

Put private files here:

```text
data/private/
```

That folder is ignored by Git.

Example:

```text
data/private/ANS-C01_real_quiz.docx
data/private/overrides_real.csv
```

Use mock files for GitHub:

```text
data/mock/network_docs.md
data/mock/mock_overrides.csv
```

## Pre-Push Scan

Run:

```powershell
python .\scripts\safety_scan.py
```

The scan is intentionally conservative. If it complains, inspect the file before
publishing.

