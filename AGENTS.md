# Codex rules — MathDocs Local

## Mission
Build offline-first Hebrew math documents for Microsoft Word/Excel on Windows.
All formulas and matrices MUST remain editable after copying into another DOCX.

## Architecture contract
- `schema.py`: strict JSON content schema; never mix content with office formatting.
- `omml.py`: native Word OMML, no rasterized equations.
- `word.py`: RTL layout, original editable DOCX.
- `excel_windows.py`: optional Excel desktop COM for interactive charts/calculation.
- `ai_optional.py`: OpenAI-only integration, explicit opt-in. Never use Claude.
- `cli.py`: entrypoint, offline by default.

## Hard requirements
1. No AI calls for formatting, matrices, chart construction, document export, or file copies.
2. No network calls during `mathdocs demo` and `mathdocs build`.
3. Preserve user originals: never overwrite a user-supplied DOCX without explicit approval.
4. No uploading document contents or keys to any service without consent.
5. Never replace mathematical expressions with pictures to hide rendering problems.
6. Use strict JSON schema validation before writing files.
7. Make a tiny change, run `python -m pytest -q`, inspect DOCX OOXML, then visually render.
8. Support Windows desktop Word/Excel first; cloud and add-ins are future phases.
9. Keep user explanations in simple Hebrew and formulas in a stable math representation.

## First Codex tasks
1. Inspect repo and run tests. Fix any errors without losing native OMML.
2. Add Word import or insertion into an existing DOCX while preserving existing sections.
3. Add an integration test opening/editing/saving math using Windows Word COM.
4. Add lightweight GUI for importing a structured lesson.
5. Add an optional OpenAI provider with token/cost metering and no default network calls.
