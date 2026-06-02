# 🚀 SheetMind — Future Scope & Feature Roadmap

Ideas for taking SheetMind beyond its current state, roughly ordered by
**impact ÷ effort**. Each item notes *what*, *why*, and *where it plugs in*.

Legend: 🟢 quick win · 🟡 medium · 🔴 large · 💲 may need paid LLM/API

---

## Tier 1 — High impact, low effort

### 🟢 Conditional formatting & cell styling
"Highlight rows where stock < 10 in red", "bold the header".
- **Where:** new `format` op in `engine.py`; apply via openpyxl styles in
  `SheetStore.save()` (web) and xlwings cell `.color` (add-in).
- **Why:** the #1 thing Excel Copilot does that we don't yet.

### 🟢 Charts from a sentence
"Make a bar chart of revenue by month".
- **Where:** `chart` read-op returning a spec; render with `st.bar_chart`/Altair in
  the web app, and an embedded Excel chart in the add-in.

### 🟢 Undo/redo + multi-step in chat
Show "Undo" inline after each applied change; add **redo**.
- **Where:** extend `SheetStore` history to a two-stack (undo/redo); UI buttons.

### 🟢 Streamlit deprecation cleanup & version pin
Migrate `use_container_width=True` → `width='stretch'`; pin `streamlit` range.
- **Why:** future-proofing; removes console warnings.

### 🟢 Command history & "re-run last"
Persist recent commands; one-click repeat / edit-and-rerun.

---

## Tier 2 — Core capability expansion

### 🟡 Multi-sheet & cross-sheet operations
"Copy matching rows to a 'Reorder' sheet", "VLOOKUP price from Sheet2".
- **Where:** engine ops gain a `sheet` target; `SheetStore` already tracks
  `sheet_names`/`active_sheet` — extend read/write to any sheet and add a `join` op.

### 🟡 Pivot tables
"Pivot sales by region and quarter".
- **Where:** `pivot` read-op using `pandas.pivot_table`; render as a table + offer
  "write to new sheet".

### 🟡 Find & fix data quality
"Remove duplicate rows", "trim whitespace", "fill blank prices with 0",
"standardize date formats".
- **Where:** `dedupe`, `fillna`, `transform` ops — pure pandas, deterministic.

### 🟡 Rich row addressing
"Update the 3rd row", "the last 5 rows", "rows 10–20".
- **Where:** extend the filter grammar with positional selectors (`index`, `head`,
  `tail`, `range`).

### 🟡 Bulk import / merge
"Append this CSV", "merge with last month's file on Item".
- **Where:** new loaders in `core.py`; a `merge` op wrapping `pandas.merge`.

---

## Tier 3 — Smarter agent

### 🟡💲 Conversational memory & follow-ups
"…now do the same for Office items", "undo that and try 20% instead".
- **Where:** keep the last N turns + last plan in the LLM context; the chat already
  stores history — feed a summary back into `interpret()`.

### 🟡 Clarifying questions UI
When `interpret()` returns an empty plan + a question, render it as a chat prompt
with quick-reply buttons instead of plain text.

### 🟡💲 "Explain this sheet" / auto-insights
One click → the model summarizes trends, outliers, and suggested actions
(grounded in deterministic stats we compute first).

### 🔴💲 Multi-step task planner
"Clean the data, then chart monthly totals, then email me the file." Decompose
into a sequence of plans with a single confirm-all / step-through.
- **Where:** a planner that emits an ordered list of plans; execute with the
  existing preview/commit per step.

### 🟡 Model router & cost controls
Auto-pick a cheap model for simple parses and a stronger one for complex asks;
show token/cost usage; cache identical (command, schema) results.
- **Where:** `llm.py` already has a model chain — add a complexity heuristic.

---

## Tier 4 — Platform & reach

### 🔴 Real-time streaming voice
Replace upload-then-transcribe with live mic streaming + partial transcripts.
- **Where:** `core.transcribe` → a streaming STT (faster-whisper locally, or a
  realtime API). Add a "push-to-talk" widget.

### 🔴 Google Sheets backend
Same engine, Google Sheets API instead of openpyxl.
- **Where:** a `GSheetStore` implementing the store interface (`.df`,
  `_push_history`, `save`).

### 🔴 Office.js web add-in
A modern task-pane add-in (works in Excel on the web/Mac), replacing the
xlwings/VBA path for broader reach.

### 🟡 Auth & multi-user (hosted)
Per-user API keys, saved sessions, file storage — for a SaaS deployment.

### 🟡 Audit log & change tracking
Every applied plan written to a log sheet / file ("who/when/what changed").

---

## Tier 5 — Quality, safety, trust

### 🟢 Expand the test suite
Add tests for the LLM path with a mocked `llm.chat`, the repair loop, and the
local parser edge cases; add CI (GitHub Actions) running `pytest`.

### 🟡 Golden-command eval harness
A fixed set of (command → expected plan) pairs run against each model to measure
accuracy and catch regressions when changing prompts/models.

### 🟢 Guardrails on destructive ops
Require typing a word (or a second confirm) for `clear_rows` / mass deletes;
cap the number of rows a single command may delete without extra confirmation.

### 🟡 Better error messages & recovery
Surface "did you mean column X?" suggestions; offer a one-click fix when a column
name is close but not exact.

### 🟢 Internationalization
The engine already accepts multilingual voice; localize the UI strings and
example hints.

---

## Suggested next sprint (if picking up tomorrow)
1. 🟢 Conditional formatting + charts (closes the biggest Copilot gap).
2. 🟢 Streamlit deprecation cleanup + version pin (stability).
3. 🟡 Conversational follow-ups (memory) — makes it *feel* like an agent.
4. 🟢 CI + mocked-LLM tests (protects everything above from regressions).
