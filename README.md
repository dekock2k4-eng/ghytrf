# 🧠 SheetMind

**Talk to your spreadsheet.** SheetMind turns plain English (typed or spoken) into
precise, previewed edits and instant analytics on your Excel files — *"raise every
electronics price by 10%"*, *"what's my total stock value?"*, *"delete rows where
quantity is below 5"*, *"add a Total column = Qty × Price"*.

It ships in **two forms** that share one brain:

| Form | What it is | Best for |
|------|------------|----------|
| 🌐 **Streamlit web app** (`app.py`) | A conversational chat UI in your browser. Upload/create a sheet, talk to it, download the result. | Trying it out, no-Excel workflows, demos |
| 🧩 **Excel ribbon add-in** (`addin.py`) | A native **SheetMind** tab inside Excel with Run / Voice / Undo / Info buttons that edit the *active sheet in place*. | Power users who live in Excel |

---

## ✨ What makes it world-class

- **Rich command vocabulary** — not just add/edit/delete. Real filters with
  operators (`>`, `<`, `contains`, `between`, `in`, …), sorting, computed/formula
  columns, multi-column updates, percentage maths, and **read-only analytics**
  (sum / average / min / max / median / count, `group by`, `top-N`) for Q&A.
- **Tool-calling agent** — the LLM never hands back loose text. It calls a strict
  `build_plan` function and returns a **validated, typed plan** of operations.
- **Self-repairing** — every plan is dry-run against your real sheet; if an
  operation is malformed or references a missing column, the exact error is fed
  back to the model to fix itself (the *structured-reflection* pattern). This is
  what lets even small **free** models hit high accuracy.
- **Deterministic maths** — all numbers in answers are computed by pandas, never
  guessed by the model.
- **Preview → confirm** — any change is shown as a diff (what changes, rows/cols
  before→after, a preview table) and only written after you approve. Plus a
  25-step undo stack and atomic saves that never corrupt your file.
- **Voice-ready & multilingual** — Whisper transcription with auto language
  detection; the engine understands commands in many languages.
- **Provider-agnostic & free by default** — runs on OpenRouter with
  **Kimi K2.6 (free)** out of the box, with automatic fallback across other free
  open models.

---

## 🏗 How it works (architecture)

```
            ┌──────────────┐        ┌──────────────┐
  voice  →  │  transcribe  │        │   your .xlsx │
 (Whisper)  │  (core.py)   │        └──────┬───────┘
            └──────┬───────┘               │ read
                   │ text            ┌──────▼───────┐
  text  ───────────┼───────────────► │  SheetStore  │  pandas DataFrame
                   │                 │  (core.py)   │  + undo + atomic save
                   ▼                 └──────┬───────┘
        ┌──────────────────────┐           │ schema + sample + stats
        │  engine.interpret()  │ ◄─────────┘
        │  ──────────────────  │
        │  1. local fast-path  │  (regex — instant, no network)
        │  2. LLM tool-call    │ ──► llm.py ──► OpenRouter (Kimi K2.6)
        │  3. validate + repair│ ◄── structured reflection loop
        └──────────┬───────────┘
                   │ Plan  (typed operations)
                   ▼
        ┌──────────────────────┐     mutation?      ┌───────────────┐
        │   engine.dry_run()   │ ──── yes ────────► │ preview + diff │ ─► confirm
        │  (executes on a copy)│                    └──────┬────────┘
        └──────────┬───────────┘     read-only             │ approve
                   │ answers/tables                         ▼
                   ▼                              ┌────────────────────┐
              show in chat                        │  engine.commit()   │
                                                  │  push undo · save  │
                                                  └────────────────────┘
```

### The pieces

| File | Responsibility |
|------|----------------|
| **`llm.py`** | Provider-abstracted LLM access. OpenAI-compatible OpenRouter client with tool calling, a configurable model chain, retries/backoff, and graceful "no key" handling. |
| **`engine.py`** | The brain. Operation executors (filters, formulas, sort, aggregate, top-N…), the `build_plan` tool schema + few-shot system prompt, the **validate→repair** loop, and `dry_run` / `commit` with preview diffs. Store-agnostic. |
| **`core.py`** | `SheetStore` — the pandas/openpyxl data layer (load, **atomic** save, 25-step undo, multi-sheet) — plus Whisper transcription and shared helpers (word-numbers, fuzzy column matching). |
| **`app.py`** | The Streamlit chat UI: Assistant (chat + preview/confirm + voice/text), Load, Data, Delete, Columns, Dashboard. |
| **`addin.py`** | The Excel add-in engine (`XlwingsStore`) and ribbon functions, using the same `engine` with Yes/No confirm dialogs. |
| **`create_addin.py`** | Builds & installs `SheetMind.xlam` into Excel. |
| **`tests/`** | `pytest` suite covering conditions, formulas, every operation, dry-run vs commit, and the local parser. |

### Why "Excel → DataFrame → operate"
Every command works on an in-memory **pandas DataFrame**, not raw cells. Numeric
columns are auto-detected on load, so `>`/`<`/`sum`/`sort` behave correctly even
when Excel stored numbers as text. The model only ever sees a compact **schema +
sample + stats** of that DataFrame (never your whole file), which keeps prompts
small, fast, cheap, and private. Results are written back atomically.

---

## 🚀 Install & run

### Prerequisites
- **Python 3.10–3.13** (3.11 recommended).
- A free **OpenRouter** API key — <https://openrouter.ai/keys>.
- *(optional, for voice)* a free **Hugging Face** token — <https://huggingface.co/settings/tokens>.

### 1) Get the code & a virtual environment
```bash
cd Sheet-Mind-main
python3.11 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2) Configure your keys
```bash
cp .env.example .env
# then edit .env and paste your OPENROUTER_API_KEY (and HF_TOKEN for voice)
```
Defaults already point at the free model `moonshotai/kimi-k2.6:free` with free
fallbacks — you only need to add your key.

### 3a) Run the web app
```bash
streamlit run app.py
```
Opens at <http://localhost:8501>. Upload an `.xlsx` (or create a blank sheet),
then go to **💬 Assistant** and start typing or talking.

### 3b) Install the Excel add-in (Windows / Excel desktop)
```bash
python create_addin.py        # or double-click install.bat on Windows
```
Then in Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse**,
select the generated `SheetMind.xlam`, tick it, **OK**. A **SheetMind** tab appears
with **Run Command**, **Voice**, **Undo**, and **Sheet Info** buttons. (The add-in
relies on the `xlwings` base add-in, which the installer sets up for you.)

> The add-in reads/writes the **active sheet in place** and keeps an in-session
> undo stack. The web app works on an uploaded/created copy you download.

---

## 💬 Using it — example commands

**Edit**
- `Increase the price of every Writing item by 15%`
- `Set Status to "Reorder" where Qty is less than 10`
- `Add 50 to Qty where Item is Pen`
- `Delete rows where Qty is below 5`
- `Add a column Total that is Qty times Price`
- `Rename column Qty to Quantity` · `Sort by Price descending`
- `Add a row: Item = Glue, Qty = 12, Price = 3`

**Ask (read-only analytics)**
- `What is the total stock value?`  → Σ (Qty × Price), computed exactly
- `Average price of Office items?`
- `Which item has the highest stock value?`
- `Top 5 rows by Qty`
- `Total quantity by Category`

Anything that changes the sheet shows a **preview** first; analytics answer
immediately. Mis-said something? **Undo** reverts the last change.

---

## ⚙️ Configuration reference (`.env`)

| Variable | Default | Purpose |
|----------|---------|---------|
| `OPENROUTER_API_KEY` | — | **Required** for AI commands. |
| `SHEETMIND_MODEL` | `moonshotai/kimi-k2.6:free` | Primary model slug. |
| `SHEETMIND_MODEL_FALLBACKS` | free deepseek + gemma-4 | Comma-separated fallbacks. |
| `HF_TOKEN` | — | Hugging Face token for Whisper voice. |
| `SHEETMIND_FORCE_LANGUAGE` | auto | Pin spoken language (`en`, `hi`, …). |
| `SHEETMIND_SAVE_DIR` | project dir | Where edited files are written. |
| `SHEETMIND_LLM_BASE_URL` | OpenRouter | Swap to another OpenAI-compatible endpoint. |
| `SHEETMIND_LLM_TIMEOUT` | `60` | Per-request timeout (seconds). |

**Want a different model?** Any tool-calling model on OpenRouter works — e.g.
`anthropic/claude-3.7-sonnet`, `openai/gpt-4o-mini`, `qwen/qwen3-235b-a22b`. Just
set `SHEETMIND_MODEL`.

---

## 🧪 Testing

```bash
pip install pytest
pytest -q
```
The suite is deterministic (no network): it feeds the executor the same plan
shape the LLM emits and checks conditions, formula maths, every operation,
dry-run-vs-commit semantics, and the local fast-path parser.

---

## 🆘 Troubleshooting

| Symptom | Fix |
|---------|-----|
| *"No OPENROUTER_API_KEY set"* | Add your key to `.env`. Simple `set/delete/add … where …` commands still work offline via the local parser. |
| Voice button errors | Set `HF_TOKEN` in `.env`; or just use **Text** mode. |
| *"save failed"* in the web app | The `.xlsx` is open in Excel — close it and retry. |
| Add-in buttons error | Ensure the **xlwings** base add-in is enabled (Excel → Add-ins). |
| Model is slow / rate-limited | Free models throttle; SheetMind auto-falls-back, or set a paid `SHEETMIND_MODEL`. |
| Wrong column picked | Use the exact header name; SheetMind also fuzzy-matches close names. |

---

## 📚 More docs
- [`docs/HANDOFF.md`](docs/HANDOFF.md) — what changed vs. the original, how it was
  built, full architecture, and current status / known issues.
- [`docs/ROADMAP.md`](docs/ROADMAP.md) — future scope: features to add next,
  ranked by impact ÷ effort.

---

## 🔒 Privacy & safety
- Only a **schema + 3 sample rows + summary stats** are sent to the model — never
  your full file.
- `.env` and all `*.xlsx` files are gitignored. **Never commit real API keys.**
- Every mutating command is previewed and confirmed; saves are atomic (write to a
  temp file then rename) so an interrupted write can't corrupt your data.
