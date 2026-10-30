# MathDocs Local — Architecture v0.1

## Goal
Create Hebrew math DOCX where each matrix/equation is native Word OMML and editable.
Use Excel desktop for calculated matrices and data-driven charts; no obligatory cloud.

```text
Notebook scans -> [ChatGPT manually or Optional OpenAI API]
               -> validated lesson JSON
               -> MathDocs Local / Python
                  |-> Word OOXML / OMML -> editable DOCX
                  |-> Excel COM (Windows) -> editable XLSX graph and matrix
                  |-> QA: OOXML tests + DOCX render + Word desktop copy/paste tests
```

## Stage 1 (implemented)
- JSON lesson contract / validation
- DOCX builder with Hebrew RTL paragraphs and OMML native math
- limit, fraction, powers, subscripts and matrices
- native table in Word
- optional Excel COM worksheet/chart demo
- OpenAI-only optional provider; default network traffic = 0
- automated structure tests and instructions for Codex

## Stage 2 (not implemented)
- insertion at cursor / bookmarks in an existing Word document
- packaged GUI for Windows
- more math structures (roots, integrals, derivatives, aligned equations)
- Windows E2E tests opening Word, copying math to a blank document, modifying it
- more Excel chart/data interop
- reversible document edits and backups

## Stage 3 (not implemented)
- Office add-in sidebar (Office.js)
- optional OpenAI understanding of handwriting / notes
- user-visible per-task token counters; budget caps; secure key storage
- cloud storage optional, not assumed

## Design decisions
- Content is JSON and independent from formatting and AI providers.
- OMML rather than screenshot / flattened LaTeX to preserve editability.
- Excel desktop COM only in logged-in Windows session; never automate GUI Office on cloud server.
- No external AI credentials needed to generate documents.
- API use explicitly opt-in; model may be changed by the user.
- Word may change visual styles when pasting into documents with conflicting styles.
  `Keep Source Formatting` helps, but true end-to-end fidelity must be tested in Word.

## Technical risks
- Word vs LibreOffice OMML rendering differences.
- RTL and LTR interactions inside the same paragraph.
- Excel COM cannot be tested outside Windows with Office installed.
- Excel data cannot become calculation-enabled in Word via simple paste.
- Complex latex-to-OMML conversion is out of scope until tested.
- AI hallucinations and handwriting ambiguity require review.

## Acceptance tests
1. DOCX `word/document.xml` contains `m:oMathPara`, `m:f`, and `m:m` nodes.
2. DOCX has zero equation screenshots or chart screenshots masquerading as editable math.
3. Equations are editable using Microsoft Word's native Equation Editor.
4. Matrix can be copied between 2 DOCX files and modified.
5. In Excel, editing matrix cells recomputes determinant, and changing x updates chart.
6. Default build works with no internet and no OPENAI_API_KEY.
7. User originals are not overwritten by default.
