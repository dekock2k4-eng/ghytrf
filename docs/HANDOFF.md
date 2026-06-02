# 🤝 SheetMind — Engineering Handoff

**Date:** 2026-06-02
**Scope of this work:** Full engine + UX revamp — turned SheetMind from a
4-command CRUD tool into a world-class, agentic natural-language spreadsheet
assistant. Nothing in the original feature set was removed; everything was
upgraded and a great deal was added.

---

## 1. TL;DR — what changed

| | **Before** | **After** |
|---|---|---|
| Command set | 4 actions: `add_row`, `set_cell`, `update_row`, `delete_row` | add_row, **update** (multi-col), delete, **clear**, **add/delete/rename column**, **sort**, plus read-only **aggregate / top_n / describe** |
| Matching | exact case-insensitive **equality only** | operators: `eq, ne, gt, gte, lt, lte, contains, startswith, endswith, between, in, isblank, notblank`, AND/OR groups |
| Maths | none | **formula columns** & set-values (`=[Qty]*[Price]*1.1`), deltas (`+5/-10`), word-numbers |
| Analytics / Q&A | ❌ none | "total stock value?", "avg price of writing items", "top 5 by qty", "total qty by category" — **computed by pandas, never guessed** |
| LLM | Gemini, single-shot, JSON scraped from a string | **OpenRouter tool-calling** (Kimi K2.6 free) with a **validate→repair loop** |
| Safety | applied instantly (undo only) | **preview diff → confirm** before any write, + 25-step undo + atomic save |
| UX | tab with voice/text buttons | **conversational chat** with answer bubbles, result tables, preview cards |
| Voice | English-only (hard-blocked other scripts) | **multilingual** auto-detect |
| Tests | ❌ none | **29 deterministic pytest tests** |
| Docs | none | `README.md`, `.env.example`, this handoff, roadmap |

---

## 2. How we did it (approach)

1. **Researched the field** (Excel Copilot, pandas-ai / LangChain DataFrame
   agents, voice-spreadsheet UX, and the *structured-reflection* paper) to define
   what "world-class" means: tool calling, schema grounding, deterministic maths,
   preview-before-apply, and self-correction.
2. **Built a clean engine layer** separate from the UIs so both the web app and
   the Excel add-in share one brain.
3. **Made the free model reliable** with three reinforcing techniques:
   - a forced `build_plan` **tool call** (structured output, not prose),
   - a **few-shot system prompt** with the exact shape of every operation,
   - a **validate→repair loop**: each plan is dry-run on the real sheet; any
     error is sent back to the model to fix itself (up to 2 rounds).
4. **Kept all maths deterministic** — the model only chooses *which* operation;
   pandas computes every number.
5. **Verified at every layer** — unit tests for the engine, headless
   `AppTest` runs for the UI preview→confirm flow, and live LLM checks.

---

## 3. Architecture (current state)

```
 voice ─(Whisper)─┐
                  ├─► engine.interpret(text, df)
 text ────────────┘        │  1. local regex fast-path (no network)
                           │  2. LLM tool-call  ── llm.py ─► OpenRouter (Kimi K2.6)
                           │  3. validate + repair (structured reflection)
                           ▼
                        Plan (typed operations)
                           │
              mutation? ───┴─── read-only
                 │                  │
        engine.dry_run()      engine.commit()  (computes answers, no save)
        → preview + diff           │
        → user confirms            ▼
                 │            show tables/values in chat
                 ▼
        engine.commit() → push undo → atomic save
```

### Files (what to read first)
| File | Role | Start here |
|------|------|-----------|
| `engine.py` | **The brain.** Operations, condition/formula evaluators, `build_plan` tool schema + prompt, validate→repair, `dry_run`/`commit`. | `interpret()`, `_exec_op()`, `_plan_with_repair()` |
| `llm.py` | OpenRouter client: model chain, tool calling, retries. | `chat()`, `model_chain()` |
| `core.py` | `SheetStore` data layer (load/atomic-save/undo/multi-sheet) + Whisper `transcribe()`. | `SheetStore`, `transcribe()` |
| `app.py` | Streamlit chat UI. | the `S().page == "assistant"` block |
| `addin.py` | Excel ribbon engine (`XlwingsStore`) + ribbon functions. | `_process_command()` |
| `tests/test_engine.py` | 29 deterministic tests. | run them |

### Key design decisions
- **Store-agnostic engine.** `engine.commit()` only needs `.df`, `_push_history()`,
  `save()`. Both `SheetStore` (file-based, web app) and `XlwingsStore` (live Excel
  sheet) satisfy this, so one engine powers both front-ends.
- **The model never sees your whole file** — only a compact schema + 3 sample rows
  + summary stats. Fast, cheap, private.
- **Formulas use `[Column]` references** and run through a restricted evaluator
  (no `__builtins__`, character-whitelisted) — safe on model/user input.
- **Preview is a real dry-run** on a copy of the DataFrame, so the diff shown is
  exactly what will be written.

---

## 4. Setup & run (current state)

### Prerequisites
- Python 3.10–3.13 (3.11 recommended)
- OpenRouter API key — https://openrouter.ai/keys (free)
- *(optional, voice)* Hugging Face token — https://huggingface.co/settings/tokens

### First-time setup
```bash
cd Sheet-Mind-main
python3.11 -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env                  # then paste your OPENROUTER_API_KEY
```
> A `.env` already exists in this working copy with a key and
> `SHEETMIND_MODEL=moonshotai/kimi-k2.6:free`. `.env` is gitignored.

### Run the web app
```bash
source .venv/bin/activate
streamlit run app.py                 # http://localhost:8501
```
Upload an `.xlsx` (or create a blank sheet) → **💬 Assistant** → type or speak.

### Install the Excel add-in (Windows / Excel desktop)
```bash
python create_addin.py               # or double-click install.bat
```
Excel → File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse →
select `SheetMind.xlam` → tick → OK. Use the **SheetMind** ribbon tab.

### Run the tests
```bash
source .venv/bin/activate
pytest -q                            # 29 tests, deterministic, no network
```

---

## 5. Current status & known issues

✅ **Working & verified**
- Engine: 29/29 unit tests pass; all files compile; clean imports.
- Live Kimi K2.6 plans correctly for hard cases (stock value, % updates via
  self-repair, top-N by computed value, filtered averages).
- Web UI preview→Apply→commit flow verified headlessly.
- Excel `XlwingsStore` works with the engine.
- Graceful degradation when the LLM is unavailable/rate-limited.

⚠️ **Known issues / caveats**
- **Free-tier rate limits.** Free OpenRouter models return `HTTP 429` after a
  modest number of calls per day. SheetMind auto-falls-back across free models and
  shows a friendly message, but for heavy use add OpenRouter credits or set a paid
  `SHEETMIND_MODEL`. Offline `set/add/delete … where …` still works via the local
  parser.
- **Streamlit deprecation.** The UI still uses `use_container_width=True` (works in
  1.58, emits a deprecation warning; will eventually need `width='stretch'`).
- **Single active sheet** at a time (multi-sheet switch exists, but commands run
  against the active sheet only).
- **Dead code left in place** intentionally for safety: `core.parse_command` /
  `apply_command` (old path) and `app._run_cmd` / `generate_examples` are unused by
  the new flow but harmless; remove during a future cleanup.

---

## 6. Quick reference — what the engine accepts

Operations (the LLM emits these as a `build_plan` call; you can also hand-build
them in tests):
```python
{"op":"add_row","values":{"Item":"Glue","Qty":12}}
{"op":"update","filter":{"conditions":[{"column":"Category","op":"eq","value":"Writing"}]},"set":{"Price":"=[Price]*1.15"}}
{"op":"delete","filter":{"conditions":[{"column":"Qty","op":"lt","value":5}]}}
{"op":"add_column","name":"Total","formula":"[Qty]*[Price]"}
{"op":"rename_column","old":"Qty","new":"Quantity"}
{"op":"sort","by":["Qty"],"ascending":false}
{"op":"aggregate","func":"sum","expr":"[Qty]*[Price]"}          # total stock value
{"op":"aggregate","func":"mean","column":"Price","group_by":"Category"}
{"op":"top_n","column":"Qty","n":5,"ascending":false}
{"op":"top_n","expr":"[Qty]*[Price]","n":1}                     # highest value item
```
See `docs/ROADMAP.md` for what to build next.
