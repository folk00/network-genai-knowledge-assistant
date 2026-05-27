# Scripts

Small demos you can run from PowerShell.

```powershell
python .\scripts\demo_local_rag.py
```

This prints:

- retrieved mock documents
- scores
- the grounded prompt that would be sent to Claude/OpenAI

GUI demo:

```powershell
python .\scripts\gui_rag_demo.py
```

The GUI is safe by default. It runs dry-run mode unless you check `Live LLM call`.

Full original GUI:

```powershell
python .\scripts\run_full_quiz_gui.py
```

This launches the richer PySide6 app from `apps/full_quiz_gui/`.
