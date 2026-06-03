# 🧩 SheetMind inside Excel

SheetMind ships **two** ways to live inside Excel. Pick one:

| | **A. Task-pane add-in** ⭐ recommended | **B. xlwings ribbon** (legacy) |
|---|---|---|
| Look | Premium web UI (dark/light auto, animations, inline diff) | Native dialogs |
| Tech | Office.js panel + local Python bridge (`server.py`) | xlwings + tkinter |
| Cross-platform | Excel on Mac, Windows, and the web | Desktop only, fiddly on Mac |
| Setup | Run one server + sideload a manifest | Build .xlam + configure xlwings |

> Want to *see* it first? `./run_addin.sh --http` then open
> <http://localhost:8765/taskpane.html?demo=1> in any browser.

---

## A. Premium task-pane add-in (recommended)

A beautiful panel docked inside Excel that reads your active sheet, previews every
change, and writes back on **Apply** — auto-themed to match Excel's light/dark mode.

### How it fits together
```
Excel ──Office.js──►  taskpane (web UI)  ──HTTP──►  server.py  ──►  engine.py ──► LLM
  ▲   reads/writes        (dark/light)      /api/plan /api/apply       (preview/commit)
  └──────────────────────  Apply writes the new grid back  ──────────────────────┘
```

### 1. Trust a localhost HTTPS cert (Office requires HTTPS)
Easiest (trusted automatically) — needs Node:
```bash
npx office-addin-dev-certs install
```
`server.py` auto-detects those certs. **No Node?** `server.py` will generate a
self-signed cert under `.certs/` on first run — trust it once: open
`https://localhost:8765` in Safari → *Show Details* → *visit this website*, or add
`.certs/localhost.crt` to Keychain Access and mark it *Always Trust*.

### 2. Start the bridge
```bash
./run_addin.sh                 # → https://localhost:8765
```
Leave it running while you use the add-in. It serves the panel **and** runs the
SheetMind engine, using the same `.env` (model + keys) as the web app.

### 3. Sideload the manifest

**macOS**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
Restart Excel → **Home** tab → **Open SheetMind** (or Insert → Add-ins →
Developer Add-ins → SheetMind).

**Windows** — share a folder, then Excel → File → Options → Trust Center → Trust
Center Settings → Trusted Add-in Catalogs → add the folder's share path → restart
→ Insert → My Add-ins → Shared Folder → SheetMind. (Or use
`npx office-addin-debugging start manifest.xml`.)

**Excel on the web** — Insert → Add-ins → Upload My Add-in → choose `manifest.xml`.

### 4. Use it
Click **Open SheetMind**. Type or speak. Ask questions (instant answers), or give
commands — you'll see a **preview with a highlighted diff** and an **Apply**
button. Done.

### Standalone / browser mode (upload + download)
You don't need Excel at all to use the panel. When it runs **outside** Excel
(`./run_addin.sh --http` → <http://localhost:8765/taskpane.html>) it shows an
**"Open a spreadsheet"** button — pick any `.xlsx`, edit it with commands, then
**Download** (the ⬇ icon top-right, or the "download .xlsx" link in the footer).
Inside Excel there's no upload/download: it reads and writes the **live active
sheet** directly, and Excel saves the workbook as usual.

### Troubleshooting (A)
| Symptom | Fix |
|---|---|
| Panel blank / "bridge offline" | Start `./run_addin.sh` and keep it running. |
| Cert warning / won't load | Run `npx office-addin-dev-certs install`, or trust `.certs/localhost.crt`. |
| Button missing | Re-copy `manifest.xml` to the `wef` folder and fully restart Excel. |
| "All models failed / 429" | LLM key/credits — see the main README. |
| Voice does nothing | Some Office hosts block the mic; type instead, or use the browser preview. |

---

## B. xlwings ribbon (legacy, native dialogs)

Still supported — see the styled tkinter dialogs in `addin.py`. Build & setup:
```bash
.venv/bin/python create_addin.py        # builds SheetMind.xlam
```
Then load `SheetMind.xlam` via Excel → Tools/Options → Excel Add-ins, point xlwings
at `.venv/bin/python` + the project folder, and (Mac) create a workbook wrapper:
`python make_wrapper.py "YourWorkbook.xlsx"`. Full notes are in the project's git
history / earlier docs. For most users, **Option A looks and feels far better.**

---

## 🔑 Configuration
Both options read the same gitignored `.env` (OPENROUTER_API_KEY, SHEETMIND_MODEL,
HF_TOKEN). Change the model once, both update.
