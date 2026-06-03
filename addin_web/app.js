/* ============================================================================
   SheetMind task pane — Office.js + bridge client
   ========================================================================== */
"use strict";

const API = "";  // same origin as this page (served by server.py)
const SAMPLE = {
  columns: ["Item", "Qty", "Price", "Category"],
  rows: [
    ["Pen", 12, 2, "Writing"], ["Marker", 20, 8, "Writing"],
    ["Eraser", 5, 3, "Office"], ["Notebook", 30, 4, "Paper"],
    ["Stapler", 8, 12, "Office"], ["Highlighter", 15, 5, "Writing"],
  ],
};

const state = { columns: [], rows: [], office: false, busy: false };

const $ = (s) => document.querySelector(s);
const timeline = $("#timeline");
const hero = $("#hero");
const cmd = $("#cmd");
const sendBtn = $("#sendBtn");
const micBtn = $("#micBtn");

/* ---- helpers ---------------------------------------------------------- */
function el(tag, cls, html) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (html != null) n.innerHTML = html;
  return n;
}
function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
function scrollEnd() { timeline.scrollTop = timeline.scrollHeight; }
function toast(msg, err) {
  const t = $("#toast");
  t.innerHTML = `<span class="tdot"></span>${esc(msg)}`;
  t.classList.toggle("err", !!err);
  t.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove("show"), 2600);
}
const SPARK = '<svg class="spark" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2l1.6 6.4L20 10l-6.4 1.6L12 18l-1.6-6.4L4 10l6.4-1.6z"/></svg>';
const AVA = '<svg viewBox="0 0 24 24"><path d="M12 2l1.6 6.4L20 10l-6.4 1.6L12 18l-1.6-6.4L4 10l6.4-1.6z"/></svg>';
const isNarrow = () => window.matchMedia("(max-width:860px)").matches;

/* ---- theme ------------------------------------------------------------ */
function setTheme(dark) {
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
}
function lum(hex) {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})/i.exec(hex || "");
  if (!m) return 0;
  const [r, g, b] = [1, 2, 3].map((i) => parseInt(m[i], 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
function applyOfficeTheme() {
  try {
    const bg = Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
    setTheme(lum(bg) < 0.5);
  } catch (e) { setTheme(true); }
}
function applyAutoTheme() {
  const f = new URLSearchParams(location.search).get("theme");
  if (f) return setTheme(f === "dark");
  setTheme(!window.matchMedia || window.matchMedia("(prefers-color-scheme: dark)").matches);
}
$("#themeToggle").onclick = () =>
  setTheme(document.documentElement.getAttribute("data-theme") !== "dark");

/* ---- view toggle (Chat / Sheet, used when stacked) -------------------- */
function setView(v) {
  document.body.setAttribute("data-view", v);
  $("#viewSeg").querySelectorAll("button").forEach((x) => x.classList.toggle("on", x.dataset.v === v));
}
$("#viewSeg").querySelectorAll("button").forEach((b) => { b.onclick = () => setView(b.dataset.v); });

/* ---- health ----------------------------------------------------------- */
async function health() {
  try {
    const j = await (await fetch(API + "/api/health")).json();
    const pretty = (j.model || "local parser only")
      .replace(/:free$/, "").split("/").pop().replace(/-/g, " ");
    $("#modelName").textContent = pretty;
    $("#modelChip").title = j.model || "local parser";
    $("#statusDot").className = "dot " + (j.configured ? "live" : "err");
  } catch (e) {
    $("#modelName").textContent = "bridge offline";
    $("#statusDot").className = "dot err";
  }
}

/* ---- sheet I/O (Office) ---------------------------------------------- */
async function readSheet() {
  try {
    await Excel.run(async (ctx) => {
      const rng = ctx.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject();
      rng.load(["values", "rowCount"]);
      await ctx.sync();
      if (rng.isNullObject || !rng.values || !rng.values.length) {
        state.columns = []; state.rows = [];
      } else {
        state.columns = (rng.values[0] || []).map((v) => String(v));
        state.rows = rng.values.slice(1);
      }
    });
  } catch (e) {
    state.columns = SAMPLE.columns; state.rows = SAMPLE.rows;
  }
  onSheetReady();
}

async function writeSheet(grid) {
  if (!state.office) {  // browser preview — just keep in memory
    state.columns = grid.columns; state.rows = grid.rows; return;
  }
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const used = ws.getUsedRangeOrNullObject();
    used.load("address"); await ctx.sync();
    if (!used.isNullObject) used.clear(Excel.ClearApplyTo.contents);
    const data = [grid.columns, ...grid.rows];
    const r = data.length, c = grid.columns.length || 1;
    ws.getRangeByIndexes(0, 0, r, c).values = data;
    await ctx.sync();
  });
  state.columns = grid.columns; state.rows = grid.rows;
}

/* ---- suggestions / sync status --------------------------------------- */
function onSheetReady() {
  const n = state.rows.length, c = state.columns.length;
  const s = $("#syncStatus");
  const src = state.office ? "synced" : (state.filename ? esc(state.filename) : "sample data");
  if (!c) {
    s.innerHTML = '<span class="sync-pulse ready"></span> no data — open a spreadsheet below';
  } else {
    s.innerHTML = `<span class="sync-pulse ready"></span> ${src} · ${n} row${n !== 1 ? "s" : ""} × ${c} col${c !== 1 ? "s" : ""}`;
  }
  renderSuggestions();
  renderSheet();
}

/* ---- live spreadsheet view ------------------------------------------- */
function renderSheet(highlight) {
  const scroll = $("#sheetScroll");
  const cols = state.columns, rows = state.rows;
  $("#sheetName").textContent = state.office ? "Active sheet"
    : (state.filename || "Sample data");
  $("#sheetMeta").textContent = cols.length ? `${rows.length} × ${cols.length}` : "";
  if (!cols.length) {
    scroll.innerHTML = `<div class="sheet-empty"><div>📄</div>` +
      `<div>${state.office ? "This sheet is empty." : "Open a spreadsheet to see it here."}</div></div>`;
    return;
  }
  const t = el("table", "grid");
  const thead = el("thead"), htr = el("tr");
  htr.appendChild(el("th", "rownum", ""));
  cols.forEach((c) => htr.appendChild(el("th", null, esc(c))));
  thead.appendChild(htr); t.appendChild(thead);
  const tb = el("tbody");
  rows.slice(0, 500).forEach((row, ri) => {
    const tr = el("tr");
    tr.appendChild(el("td", "rownum", ri + 1));
    cols.forEach((_, ci) => {
      const hl = highlight && highlight.has(ri + "," + ci);
      tr.appendChild(el("td", hl ? "flash" : null, esc(row[ci])));
    });
    tb.appendChild(tr);
  });
  t.appendChild(tb);
  scroll.innerHTML = ""; scroll.appendChild(t);
}

function diffCells(oldRows, newRows) {
  const set = new Set();
  for (let i = 0; i < newRows.length; i++) {
    const o = oldRows[i];
    if (!o) { for (let j = 0; j < newRows[i].length; j++) set.add(i + "," + j); continue; }
    for (let j = 0; j < newRows[i].length; j++)
      if (String(o[j]) !== String(newRows[i][j])) set.add(i + "," + j);
  }
  return set;
}
function renderSuggestions() {
  const box = $("#suggestions"); box.innerHTML = "";
  const cols = state.columns;
  let ideas;
  if (cols.length >= 2) {
    const b = cols[1], cc = cols[2] || cols[1];
    ideas = [
      `What's the total <b>${esc(b)}</b>?`,
      `Show top 5 rows by <b>${esc(b)}</b>`,
      `Increase <b>${esc(b)}</b> by 10%`,
      `Add a column <b>Total</b> = ${esc(b)} × ${esc(cc)}`,
    ];
  } else {
    ideas = ["Summarize this sheet", "Add a row", "Sort by the first column", "How many rows are there?"];
  }
  ideas.forEach((html, i) => {
    const b = el("button", "sugg", html);
    b.style.animationDelay = (i * 70 + 120) + "ms";
    b.onclick = () => { cmd.value = b.textContent; autoGrow(); send(); };
    box.appendChild(b);
  });
}

/* ---- message rendering ----------------------------------------------- */
function hideHero() { if (hero) hero.style.display = "none"; }
function showHero() { if (hero) hero.style.display = ""; }
function addUser(text) {
  hideHero();
  const m = el("div", "msg user");
  m.appendChild(el("div", "bubble", esc(text)));
  timeline.appendChild(m); scrollEnd();
}
function addBot(innerEl) {
  const m = el("div", "msg bot");
  m.appendChild(el("div", "avatar", AVA));
  const b = el("div", "bubble");
  b.appendChild(innerEl);
  m.appendChild(b);
  timeline.appendChild(m); scrollEnd();
  return { msg: m, bubble: b };
}
function thinkingEl() {
  return el("div", "thinking",
    `<span class="bar"></span><span class="dots"><span>·</span><span>·</span><span>·</span></span>`);
}
function tableEl(grid, diff) {
  const wrap = el("div", "tbl-wrap");
  const t = el("table", "tbl");
  const thead = el("thead");
  const htr = el("tr");
  grid.columns.forEach((c) => htr.appendChild(el("th", null, esc(c))));
  thead.appendChild(htr); t.appendChild(thead);
  const tb = el("tbody");
  grid.rows.slice(0, 60).forEach((row, ri) => {
    const tr = el("tr");
    const status = diff && diff.row_status ? diff.row_status[ri] : null;
    if (status === "added") tr.className = "r-added";
    const changed = diff && diff.cell_changes ? (diff.cell_changes[ri] || []) : [];
    row.forEach((cell, ci) => {
      const td = el("td", changed.includes(ci) ? "c-changed" : null, esc(cell));
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });
  t.appendChild(tb); wrap.appendChild(t);
  return wrap;
}

/* ---- query / message results ----------------------------------------- */
function renderResultSteps(bubble, steps) {
  bubble.innerHTML = `<div class="who">SheetMind</div>`;
  let any = false;
  steps.forEach((s) => {
    if (s.error) {
      bubble.appendChild(el("div", null, `⚠ ${esc(s.summary)}: ${esc(s.error)}`)); any = true;
    } else if (s.value !== null && s.value !== undefined) {
      const m = el("div", "metric");
      m.appendChild(el("div", "label", esc(s.summary)));
      m.appendChild(el("div", "value", esc(s.value)));
      bubble.appendChild(m); any = true;
    } else if (s.table) {
      bubble.appendChild(el("div", null, `<span style="color:var(--sub);font-size:12px">${esc(s.summary)}</span>`));
      bubble.appendChild(tableEl(s.table)); any = true;
    } else {
      bubble.appendChild(el("div", null, "✓ " + esc(s.summary))); any = true;
    }
  });
  if (!any) bubble.appendChild(el("div", null, "Done."));
  scrollEnd();
}

/* ---- preview card (mutation review) ---------------------------------- */
function renderPreview(bubble, resp) {
  const card = el("div", "preview");
  card.innerHTML =
    `<div class="preview-head"><span class="tag">${SPARK} proposed change</span></div>`;
  if (resp.summary) card.appendChild(el("div", "preview-summary", esc(resp.summary)));

  const steps = el("div", "steps");
  resp.steps.forEach((s) => {
    const row = el("div", "step",
      `<span class="b"></span><span>${esc(s.summary)}${s.affected ? ` <span class="aff">(${s.affected} affected)</span>` : ""}</span>`);
    steps.appendChild(row);
  });
  card.appendChild(steps);

  if (resp.preview) {
    const removed = resp.diff && resp.diff.removed ? resp.diff.removed : 0;
    card.appendChild(el("div", "diffbar",
      `<span>preview · <b>${resp.preview.rows.length}</b> rows × <b>${resp.preview.columns.length}</b> cols</span>` +
      (removed ? `<span style="color:var(--danger)">−${removed} removed</span>` : "")));
    card.appendChild(tableEl(resp.preview, resp.diff));
  }

  const actions = el("div", "preview-actions");
  const apply = el("button", "btn apply",
    `<svg viewBox="0 0 24 24"><path d="M5 13l4 4L19 7"/></svg> Apply`);
  const cancel = el("button", "btn cancel",
    `<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18"/></svg> Cancel`);
  actions.append(apply, cancel);
  card.appendChild(actions);
  bubble.innerHTML = `<div class="who">SheetMind</div>`;
  bubble.appendChild(card);
  scrollEnd();

  cancel.onclick = () => {
    card.classList.add("done");
    bubble.appendChild(el("div", null, `<span style="color:var(--dim);font-size:12px">Cancelled — nothing changed.</span>`));
  };
  apply.onclick = async () => {
    apply.disabled = true; apply.innerHTML = "Applying…";
    try {
      const j = await postJSON("/api/apply", {
        operations: resp.operations, summary: resp.summary,
        columns: state.columns, rows: state.rows,
      });
      if (!j.ok) throw new Error(j.error || "apply failed");
      const oldRows = state.rows;
      await writeSheet(j.data);
      renderSheet(diffCells(oldRows, state.rows));   // show + flash the changes
      card.classList.add("done");
      const okline = el("div", null, `<span style="color:var(--accent);font-size:12px">✓ Applied to the sheet</span>`);
      bubble.appendChild(okline);
      toast("Applied to your sheet");
      if (isNarrow()) setView("sheet");              // reveal the result on small panes
    } catch (e) {
      apply.disabled = false; apply.innerHTML = "Apply";
      toast(e.message || "Apply failed", true);
    }
    scrollEnd();
  };
}

/* ---- networking ------------------------------------------------------- */
async function postJSON(path, body) {
  const r = await fetch(API + path, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return r.json();
}

/* ---- main send -------------------------------------------------------- */
async function send() {
  const text = cmd.value.trim();
  if (!text || state.busy) return;
  state.busy = true; sendBtn.disabled = true;
  addUser(text);
  cmd.value = ""; autoGrow();
  const { bubble } = addBot(thinkingEl());

  try {
    const resp = await postJSON("/api/plan", {
      command: text, columns: state.columns, rows: state.rows,
    });
    if (!resp.ok) {
      bubble.parentElement.classList.add("bot");
      bubble.className = "bubble err";
      bubble.innerHTML = `<div class="who">SheetMind</div>⚠ ${esc(resp.error || "Something went wrong")}`;
    } else if (resp.kind === "message" || !resp.operations.length) {
      bubble.innerHTML = `<div class="who">SheetMind</div>${esc(resp.summary)}`;
    } else if (resp.needs_confirm) {
      renderPreview(bubble, resp);
    } else {
      renderResultSteps(bubble, resp.steps);  // read-only Q&A
    }
  } catch (e) {
    bubble.className = "bubble err";
    bubble.innerHTML = `<div class="who">SheetMind</div>⚠ ${esc(e.message || "Bridge error")}`;
  } finally {
    state.busy = false; sendBtn.disabled = false; scrollEnd();
  }
}

/* ---- composer behaviour ---------------------------------------------- */
function autoGrow() {
  cmd.style.height = "auto";
  cmd.style.height = Math.min(cmd.scrollHeight, 120) + "px";
}
cmd.addEventListener("input", autoGrow);
cmd.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); }
});
sendBtn.onclick = send;
$("#resyncBtn").onclick = () => { state.office ? readSheet() : onSheetReady(); toast("Re-synced"); };

/* ---- voice ------------------------------------------------------------ */
let rec, chunks = [];
micBtn.onclick = async () => {
  if (rec && rec.state === "recording") { rec.stop(); return; }
  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    chunks = []; rec = new MediaRecorder(stream);
    rec.ondataavailable = (e) => chunks.push(e.data);
    rec.onstop = async () => {
      stream.getTracks().forEach((t) => t.stop());
      micBtn.classList.remove("recording");
      toast("Transcribing…");
      try {
        const blob = new Blob(chunks, { type: "audio/webm" });
        const r = await fetch(API + "/api/transcribe", {
          method: "POST", headers: { "Content-Type": "application/octet-stream" }, body: blob,
        });
        const j = await r.json();
        if (j.ok) { cmd.value = j.text; autoGrow(); cmd.focus(); }
        else toast(j.error || "Transcription failed", true);
      } catch (e) { toast("Transcription failed", true); }
    };
    rec.start(); micBtn.classList.add("recording"); toast("Listening… tap mic to stop");
  } catch (e) { toast("Microphone unavailable here", true); }
};

/* ---- upload / download (standalone, when not inside Excel) ------------- */
async function loadFile(file) {
  if (!file) return;
  toast("Reading " + file.name + "…");
  try {
    const buf = await file.arrayBuffer();
    const r = await fetch(API + "/api/load", {
      method: "POST", headers: { "Content-Type": "application/octet-stream" }, body: buf,
    });
    const j = await r.json();
    if (!j.ok) { toast(j.error || "Could not read file", true); return; }
    state.columns = j.columns; state.rows = j.rows; state.filename = file.name;
    document.querySelectorAll(".msg").forEach((n) => n.remove());
    showHero(); onSheetReady();
    toast("Loaded " + file.name);
    if (isNarrow()) setView("sheet");   // let them see it loaded
  } catch (e) { toast("Could not read file", true); }
}

async function downloadFile() {
  if (!state.columns.length) { toast("Nothing to download yet", true); return; }
  try {
    const r = await fetch(API + "/api/save", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ columns: state.columns, rows: state.rows,
                             filename: state.filename || "SheetMind.xlsx" }),
    });
    const blob = await r.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = state.filename || "SheetMind.xlsx";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    toast("Downloaded " + a.download);
  } catch (e) { toast("Download failed", true); }
}

function setupStandaloneControls() {
  // hidden file picker
  const fi = el("input"); fi.type = "file"; fi.accept = ".xlsx,.xls"; fi.style.display = "none";
  fi.onchange = () => { if (fi.files[0]) loadFile(fi.files[0]); fi.value = ""; };
  document.body.appendChild(fi);

  // hero "Open" button (empty-state CTA)
  const open = el("button", "open-btn",
    `<svg viewBox="0 0 24 24"><path d="M12 16V4M7 9l5-5 5 5M5 20h14"/></svg> Open a spreadsheet`);
  open.onclick = () => fi.click();
  hero.insertBefore(open, $("#suggestions"));

  // sheet-pane toolbar: Open + Download (always reachable beside the grid)
  const tools = $("#paneTools");
  const openT = el("button", "tool-btn",
    `<svg viewBox="0 0 24 24"><path d="M12 16V4M7 9l5-5 5 5M5 20h14"/></svg> Open`);
  openT.onclick = () => fi.click();
  const dlT = el("button", "tool-btn accent",
    `<svg viewBox="0 0 24 24"><path d="M12 4v12M7 11l5 5 5-5M5 20h14"/></svg> Download`);
  dlT.onclick = downloadFile;
  tools.append(openT, dlT);

  // footer "re-sync" → download (re-sync only matters inside Excel)
  const rs = $("#resyncBtn");
  if (rs) { rs.textContent = "download .xlsx"; rs.onclick = downloadFile; }
}

/* ---- boot ------------------------------------------------------------- */
let booted = false;
function bootBrowser() {
  if (booted) return; booted = true;
  state.office = false; applyAutoTheme();
  state.columns = SAMPLE.columns; state.rows = SAMPLE.rows; state.filename = null;
  setupStandaloneControls();
  onSheetReady();
}
/* demo seed — `?demo=1` shows a sample conversation (handy for previews/screens) */
function seedDemo() {
  hideHero();
  addUser("What's the total stock value?");
  const q = addBot(el("div"));
  renderResultSteps(q.bubble, [{ summary: "sum([Qty]*[Price]) where all rows", value: 318, kind: "read" }]);
  addUser("Raise every Writing item's price by 10%, rounded");
  const p = addBot(el("div"));
  renderPreview(p.bubble, {
    summary: "Increase price of Writing items by 10% and round to the nearest integer.",
    operations: [{ op: "update" }],
    steps: [{ summary: "Set Price = ROUND([Price]×1.1) where Category = Writing", affected: 3 }],
    preview: {
      columns: ["Item", "Qty", "Price", "Category"],
      rows: [["Pen", 12, 2, "Writing"], ["Marker", 20, 9, "Writing"],
             ["Eraser", 5, 3, "Office"], ["Notebook", 30, 4, "Paper"],
             ["Highlighter", 15, 6, "Writing"]],
    },
    diff: { row_status: ["changed", "changed", "normal", "normal", "changed"],
            cell_changes: [[2], [2], [], [], [2]], removed: 0 },
  });
}

health();
if (new URLSearchParams(location.search).has("demo")) {
  booted = true; applyAutoTheme();
  state.columns = SAMPLE.columns; state.rows = SAMPLE.rows; state.filename = "Inventory.xlsx";
  setupStandaloneControls();
  renderSheet(); renderSuggestions();
  setTimeout(seedDemo, 200);
} else if (window.Office && Office.onReady) {
  Office.onReady((info) => {
    if (info && info.host) {
      booted = true; state.office = true;
      applyOfficeTheme();
      try {
        Office.context.officeTheme && Office.addin && Office.addin.onVisibilityModeChanged;
      } catch (e) {}
      readSheet();
    } else {
      bootBrowser();
    }
  });
  setTimeout(bootBrowser, 1500);  // browser preview fallback
} else {
  bootBrowser();
}

// optional initial view for previews: ?view=sheet|chat
const _vp = new URLSearchParams(location.search).get("view");
if (_vp) setTimeout(() => setView(_vp), 60);
