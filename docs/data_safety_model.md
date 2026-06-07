# Data Safety Model

This repository is designed to show the architecture without publishing private
or protected material.

## Safe To Publish

The repo may include:

- application source code
- synthetic mock data
- README and documentation
- safety scan script
- test files
- `.env.example`
- empty/private-folder README files

## Not Safe To Publish

Do not commit:

- real certification question banks
- dumps or answer keys
- override CSVs based on protected banks
- generated AI reports based on private question text
- concept dossiers generated from protected content
- customer configs
- production logs
- credentials
- API keys
- local Claude/OpenAI auth/session material
- screenshots with protected data

## Git Ignore Strategy

The `.gitignore` blocks:

- `.env`
- logs
- caches
- virtual environments
- private data folders
- DOCX/XLSX/PDF files
- generated reports
- concept dossiers
- override CSVs
- private key files

`data/private/README.md` is kept so the folder exists, but private files inside
that folder remain ignored.

## Safety Scan

Run:

```powershell
python .\scripts\safety_scan.py
```

The scan is intentionally simple. It checks for obvious private file types,
secrets and risky artifacts before publishing. It does not replace careful human
review.

## Recommended Publishing Checklist

Before pushing:

1. Run `git status --short --ignored`.
2. Confirm logs and caches are ignored.
3. Run `python .\scripts\safety_scan.py`.
4. Run tests.
5. Review staged files with `git diff --cached --name-only`.
6. Do not stage private DOCX/CSV/report files.

## Summary

The public repo is sanitized. It includes the code, mock data and architecture
docs, but excludes real question banks, protected material, generated reports
and credentials. GenAI projects need data boundaries from the beginning.
