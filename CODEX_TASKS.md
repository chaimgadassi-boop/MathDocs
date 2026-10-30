# Codex handoff

After extracting archive:

```powershell
cd MathDocs-starter
py -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
python -m pip install -e ".[dev,windows]"
python -m pytest -q
python -m mathdocs demo --output output\\example.docx
codex
```

Prompt Codex in the project root:
> Read AGENTS.md and docs/ARCHITECTURE.md. Run tests, review native OMML correctness, improve compatibility with Word for Windows, then implement a safe insert-into-existing-document command with backups. Never call OpenAI API by default; keep editable math in OMML.

## Quick manual QA
- Open `output/example.docx` in Microsoft Word desktop.
- Click inside the example limit and edit a coefficient.
- Click inside matrix, edit the central 4 to 9.
- Copy/paste each into a fresh Word document; choose Keep Source Formatting.
- Run `python -m mathdocs excel-demo --output output\\matrix_chart.xlsx` after installing Excel desktop and pywin32.
- Change B3 in Excel; observe MDETERM and graph updates.
- Report screenshots of discrepancies, never assume that screenshot appearance proves math is editable.
