"""
SheetMind  —  app.py
Run:  streamlit run app.py
"""
from __future__ import annotations
import streamlit as st
import pandas as pd
import os
from datetime import datetime
from dotenv import load_dotenv



load_dotenv()

st.set_page_config(
    page_title="SheetMind",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Inter:wght@400;500;600;700&display=swap');
:root{
  --bg:#0b1220; --s1:#111827; --s2:#0f172a; --s3:#111827;
  --b1:#1f2937; --b2:#334155;
  --blue:#38bdf8; --grn:#22c55e; --red:#f87171; --amb:#fbbf24; --pur:#a78bfa;
  --txt:#e2e8f0; --sub:#94a3b8; --dim:#64748b;
  --mono:'JetBrains Mono',monospace; --sans:'Inter',sans-serif;
}
html,body,[data-testid="stAppViewContainer"],[data-testid="stMain"],section.main{
  background:var(--bg)!important;font-family:var(--sans);color:var(--txt);}
[data-testid="stSidebar"],[data-testid="stDecoration"],
[data-testid="stHeader"],[data-testid="stToolbar"]{display:none!important;}

/* ── Nav ── */
.nav-lbl{font-family:var(--mono);font-size:.72rem;font-weight:600;
  letter-spacing:.1em;text-transform:uppercase;padding:10px 14px;
  border-radius:999px;color:var(--sub);background:rgba(148,163,184,.08);
  display:block;text-align:center;transition:.18s ease;}
.nav-lbl.on{color:var(--blue);background:rgba(56,189,248,.16);border:1px solid rgba(56,189,248,.18);}

/* ── Cards ── */
.card{background:var(--s2);border:1px solid var(--b1);
  border-radius:18px;padding:22px 24px;margin-bottom:16px;
  box-shadow:0 18px 45px rgba(0,0,0,.12);}
.clabel{font-family:var(--mono);font-size:.68rem;font-weight:700;
  letter-spacing:.12em;text-transform:uppercase;color:var(--sub);
  display:block;margin-bottom:14px;}

/* ── Buttons ── */
.stButton>button{font-family:var(--mono)!important;font-size:.85rem!important;
  font-weight:600!important;letter-spacing:.04em!important;
  border-radius:14px!important;border:1px solid var(--b2)!important;
  background:var(--s1)!important;color:var(--txt)!important;
  padding:12px 18px!important;transition:all .16s!important;width:100%;}
.stButton>button:hover{border-color:var(--blue)!important;
  color:#fff!important;background:rgba(56,189,248,.14)!important;}
[data-testid="baseButton-primary"]{background:var(--blue)!important;
  color:#0f172a!important;border-color:var(--blue)!important;}
[data-testid="baseButton-primary"]:hover{background:#22c55e!important;color:#0f172a!important;}
.rb .stButton>button{border-color:var(--red)!important;color:var(--red)!important;}
.rb .stButton>button:hover{background:rgba(248,113,113,.12)!important;}
.gb .stButton>button{border-color:var(--grn)!important;color:var(--grn)!important;}
.gb .stButton>button:hover{background:rgba(34,197,94,.12)!important;}

/* ── Pills ── */
.pill{display:inline-flex;align-items:center;font-family:var(--mono);
  font-size:.67rem;padding:5px 11px;border-radius:999px;font-weight:600;
  background:rgba(59,130,246,.12);color:var(--blue;}
.pg{background:rgba(34,197,94,.12);color:var(--grn);}
.pb{background:rgba(56,189,248,.12);color:var(--blue);}
.pa{background:rgba(251,191,36,.12);color:var(--amb);}
.pp{background:rgba(167,139,250,.12);color:var(--pur);}
.pr{background:rgba(248,113,113,.12);color:var(--red);}

/* ── Command result boxes ── */
.tx{background:rgba(56,189,248,.08);border:1px solid rgba(56,189,248,.18);
  border-left:4px solid var(--blue);border-radius:14px;padding:16px 18px;
  font-family:var(--mono);font-size:.9rem;color:#cfe8ff;line-height:1.7;margin:12px 0;}
.rx-ok{background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.2);
  border-left:4px solid var(--grn);border-radius:14px;padding:14px 16px;
  font-family:var(--mono);font-size:.82rem;color:#bbf7d0;margin:8px 0;}
.rx-warn{background:rgba(251,191,36,.08);border:1px solid rgba(251,191,36,.2);
  border-left:4px solid var(--amb);border-radius:14px;padding:14px 16px;
  font-family:var(--mono);font-size:.82rem;color:#fde68a;margin:8px 0;}

/* ── Activity log ── */
.logwrap{max-height:180px;overflow-y:auto;}
.lrow{display:flex;gap:10px;padding:6px 0;border-bottom:1px solid rgba(148,163,184,.12);
  font-family:var(--mono);font-size:.72rem;}
.lt{color:var(--dim);min-width:60px;flex-shrink:0;}
.lok{color:var(--grn);}.ler{color:var(--red);}.lin{color:var(--blue);}

/* ── Inputs ── */
.stTextInput>div>div>input,.stTextArea>div>div>textarea{
  background:rgba(15,23,42,.95)!important;border:1px solid var(--b2)!important;
  border-radius:14px!important;font-family:var(--mono)!important;
  font-size:.92rem!important;color:var(--txt)!important;padding:12px!important;}
.stTextInput>div>div>input:focus,.stTextArea>div>div>textarea:focus{
  border-color:var(--blue)!important;
  box-shadow:0 0 0 3px rgba(56,189,248,.12)!important;}
label{color:var(--sub)!important;font-size:.82rem!important;font-family:var(--sans)!important;}

/* ── File uploader ── */
[data-testid="stFileUploaderDropzone"]{background:rgba(15,23,42,.95)!important;
  border:1.5px dashed rgba(148,163,184,.22)!important;border-radius:18px!important;}
[data-testid="stFileUploaderDropzone"]:hover{border-color:var(--blue)!important;}
[data-testid="stFileUploaderDropzone"] p{color:var(--sub)!important;}

/* ── Dataframe ── */
[data-testid="stDataFrame"]{border-radius:18px;overflow:hidden;}
[data-testid="stDataFrame"] thead tr th{background:rgba(15,23,42,.95)!important;
  font-family:var(--mono)!important;font-size:.72rem!important;
  text-transform:uppercase!important;letter-spacing:.08em!important;color:var(--sub)!important;}
[data-testid="stDataFrame"] tbody tr td{font-family:var(--mono)!important;
  font-size:.82rem!important;color:var(--txt)!important;}
[data-testid="stDataFrame"] tbody tr:hover td{background:rgba(148,163,184,.08)!important;}

/* ── Metrics ── */
[data-testid="stMetric"]{background:var(--s1);border:1px solid var(--b1);
  border-radius:16px;padding:14px 16px;}
[data-testid="stMetric"] label{font-family:var(--mono)!important;
  font-size:.62rem!important;text-transform:uppercase!important;letter-spacing:.12em!important;}
[data-testid="stMetric"] [data-testid="stMetricValue"]{
  font-family:var(--mono)!important;font-weight:700!important;
  font-size:1.4rem!important;color:var(--txt)!important;}

/* ── Selectbox ── */
.stSelectbox>div>div{background:rgba(15,23,42,.95)!important;
  border:1px solid var(--b2)!important;border-radius:14px!important;
  font-family:var(--mono)!important;font-size:.9rem!important;color:var(--txt)!important;}

/* ── How-it-works ── */
.howbox{background:rgba(15,23,42,.95);border:1px dashed rgba(56,189,248,.16);
  border-radius:18px;padding:18px 20px;margin:10px 0;}
.hstep{display:flex;gap:10px;align-items:flex-start;margin-bottom:10px;}
.hnum{width:24px;height:24px;border-radius:50%;flex-shrink:0;
  background:rgba(56,189,248,.12);border:1px solid rgba(56,189,248,.22);
  font-family:var(--mono);font-size:.7rem;color:var(--blue);
  display:flex;align-items:center;justify-content:center;font-weight:700;}
.htxt{font-family:var(--mono);font-size:.78rem;color:var(--sub);line-height:1.6;}
.htxt b{color:var(--txt);}

/* ── Col list ── */
.col-row{font-family:var(--mono);font-size:.78rem;padding:8px 0;
  border-bottom:1px solid rgba(148,163,184,.12);color:var(--txt);}
.col-idx{color:var(--dim);margin-right:10px;}

/* ── Misc ── */
hr{border-color:rgba(148,163,184,.12)!important;margin:10px 0!important;}
.stCheckbox label{font-family:var(--mono)!important;font-size:.78rem!important;}
[data-testid="stAlert"]{border-radius:16px!important;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════
for k, v in {
    "store":        None,
    "page":         "load",
    "log":          [],
    "transcript":   "",
    "audio_bytes":  None,
    "last_result":  "",
    "input_mode":   "voice",   # "voice" | "text"
    "text_cmd":     "",        # typed command buffer
    "chat":         [],        # conversational history (list of message dicts)
    "pending":      None,      # PreviewResult awaiting confirm
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

def S():      return st.session_state
def store():  return S().store
def loaded(): return store() is not None and store().is_loaded

def go(page: str) -> None:
    S().page = page
    st.rerun()

def add_log(msg: str, kind: str = "info") -> None:
    S().log.insert(0, {"t": datetime.now().strftime("%H:%M:%S"),
                        "m": str(msg)[:120], "k": kind})
    if len(S().log) > 50:
        S().log.pop()

def _run_cmd(text: str) -> None:
    """Parse + apply + save a command, update last_result and log."""
    from core import parse_command, apply_command
    with st.spinner("Parsing command…"):
        try:
            action, err = parse_command(text, store())
            if err:
                add_log(err, "err")
                st.error(err)
                return
        except Exception as e:
            add_log(str(e), "err")
            st.error(str(e))
            return

    with st.spinner("Updating sheet…"):
        try:
            result = apply_command(action, store())
            S().last_result = result
            add_log(result, "ok")
            # save() is called inside apply_command → SheetStore CRUD methods.
            # No extra call needed. Just rerun to refresh the UI.
            st.rerun()
        except RuntimeError as e:
            # RuntimeError means save() failed — file could not be written
            err_msg = str(e)
            add_log(f"SAVE FAILED: {err_msg}", "err")
            st.error(
                "⚠️ Command ran but **save failed**: "
                + err_msg
                + "\n\nCheck that the .xlsx file is not open in Excel, "
                "and that the folder has write permissions."
            )
        except Exception as e:
            add_log(str(e), "err")
            st.error(str(e))


# ══════════════════════════════════════════════════════════════════
# TOP NAV BAR
# ══════════════════════════════════════════════════════════════════
NAV = [
    ("load",      "📂 Load"),
    ("assistant", "💬 Assistant"),
    ("data",      "📋 Data"),
    ("delete",    "🗑 Delete"),
    ("columns",   "🔧 Columns"),
    ("dashboard", "📊 Dashboard"),
]

brand_col, nav_col = st.columns([1, 5])

with brand_col:
    badge = (f'<span class="pill pb" style="font-size:.58rem;">'
             f'📄 {store().file_name}</span>') if loaded() else ""
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:9px;padding:5px 0 0;">
      <div style="width:26px;height:26px;border-radius:6px;flex-shrink:0;
        background:linear-gradient(135deg,#3b82f6,#8b5cf6);
        display:flex;align-items:center;justify-content:center;font-size:.78rem;">🧠</div>
      <div>
        <div style="font-size:.88rem;font-weight:700;color:#dde6f5;letter-spacing:-.02em;">
          SheetMind</div>
        {badge}
      </div>
    </div>""", unsafe_allow_html=True)

with nav_col:
    nav_cols = st.columns(len(NAV))
    for i, (pg, label) in enumerate(NAV):
        with nav_cols[i]:
            if S().page == pg:
                st.markdown(f'<div class="nav-lbl on">{label}</div>',
                            unsafe_allow_html=True)
            else:
                if st.button(label, key=f"nav_{pg}", use_container_width=True):
                    go(pg)

st.markdown('<hr style="margin:0 0 16px!important;">', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# SHARED HELPERS
# ══════════════════════════════════════════════════════════════════
def guard() -> bool:
    if not loaded():
        st.markdown("""
        <div style="text-align:center;padding:60px 0 40px;">
          <div style="font-size:2rem;margin-bottom:10px;">📂</div>
          <div style="font-family:var(--mono);font-size:.8rem;color:var(--dim);">
            No file loaded — go to
            <b style="color:var(--blue);">Load File</b> first
          </div>
        </div>""", unsafe_allow_html=True)
        if st.button("→  Go to Load File", type="primary"):
            go("load")
        return True
    return False

# ── LIVE PREVIEW ─────────────────────────────────────────────────
def render_preview(height=260):
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(
        '<span class="clabel">📄 Live Preview</span>',
        unsafe_allow_html=True
    )

    s = store()   # ✅ FIRST

    if s and s.is_loaded:


        df = s.df

        row_h = min(height, max(100, 38 + len(df) * 35))

        st.dataframe(
            df,
            use_container_width=True,
            height=row_h,
            hide_index=True
        )

    else:
        st.markdown(
            '<span class="pill pa">No file loaded</span>',
            unsafe_allow_html=True
        )

    st.markdown('</div>', unsafe_allow_html=True)


def render_log() -> None:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<span class="clabel">🕒 Activity</span>', unsafe_allow_html=True)
    if not S().log:
        st.markdown('<span style="font-family:var(--mono);font-size:.68rem;'
                    'color:var(--dim);">Nothing yet…</span>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="logwrap">', unsafe_allow_html=True)
        for e in S().log[:12]:
            cls  = {"ok":"lok","err":"ler","info":"lin"}.get(e["k"],"lin")
            icon = {"ok":"✓","err":"✗","info":"·"}.get(e["k"],"·")
            st.markdown(
                f'<div class="lrow"><span class="lt">{e["t"]}</span>'
                f'<span class="{cls}">{icon} {e["m"]}</span></div>',
                unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
def generate_examples(df):
    cols = df.columns.tolist()

    if not cols:
        return ["Upload a sheet to see examples"]

    first = cols[0]

    examples = []

    if len(cols) >= 2:
        examples.append(
            f'Add a new row with {first}=Example and {cols[1]}=Value'
        )

    if len(cols) >= 2:
        examples.append(
            f'Set {cols[1]} to NewValue where {first} is Example'
        )

    examples.append(
        f'Delete row where {first} is Example'
    )

    return examples

def dl_button(label: str = "⬇️  Download .xlsx") -> None:
    s = store()
    if s and s.file_bytes:
        st.download_button(label, data=s.file_bytes, file_name=s.file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# PAGE: LOAD FILE
# Two-column layout: LEFT = upload/create, RIGHT = preview + how-it-works
# Auto-redirects to Voice & AI as soon as a file is uploaded
# ══════════════════════════════════════════════════════════════════
if S().page == "load":
    from core import SheetStore

    left, right = st.columns([1, 1], gap="large")

    # ── LEFT: upload / create ──────────────────────────────────────
    with left:

        if loaded():
            # File already in memory — show status and CTA
            s = store()
            st.markdown(f"""
            <div style="background:rgba(16,185,129,.08);border:1px solid rgba(16,185,129,.22);
              border-radius:10px;padding:14px 18px;margin-bottom:14px;
              display:flex;align-items:center;gap:12px;">
              <div style="font-size:1.4rem;flex-shrink:0;">✅</div>
              <div>
                <div style="font-size:.88rem;font-weight:600;color:#6ee7b7;margin-bottom:2px;">
                  {s.file_name}</div>
                <div style="font-family:var(--mono);font-size:.64rem;color:var(--dim);">
                  {len(s.df)} rows · {len(s.df.columns)} cols · {s.active_sheet}</div>
              </div>
            </div>""", unsafe_allow_html=True)
            if st.button("💬  Go to Assistant  →",
                         type="primary", use_container_width=True):
                go("assistant")
            st.markdown('<div style="font-family:var(--mono);font-size:.58rem;'
                        'color:var(--dim);text-align:center;letter-spacing:.08em;'
                        'text-transform:uppercase;margin:12px 0 6px;">'
                        'Or upload a different file</div>', unsafe_allow_html=True)
        else:
            examples = ["Upload an Excel file to see examples"]
            html = '<div style="font-family:var(--mono);font-size:.65rem;color:var(--dim);line-height:1.8;margin-bottom:12px;">'
            for ex in examples:
                html += f'<span style="color:#93c5fd;">"{ex}"</span><br>'
            html += '</div>'
            st.markdown(html, unsafe_allow_html=True)

        # Upload widget
        uploaded = st.file_uploader(
            "Drop .xlsx here, or click to browse",
            type=["xlsx", "xls"], key="fu",
            label_visibility="visible")

        if uploaded:
            raw = uploaded.read()
            try:
                ns = SheetStore().load_from_upload(raw, uploaded.name)
                S().store = ns
                S().transcript = ""
                S().last_result = ""
                S().audio_bytes = None
                S().text_cmd = ""
                add_log(
                    f"Loaded: {uploaded.name} "
                    f"({len(ns.df)} rows, {len(ns.df.columns)} cols)",
                    "ok"
                )
                go("assistant")
            except Exception as e:
                st.error(f"Failed to load: {e}")
                add_log(str(e), "err")

        elif not loaded():
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">📝 Create a Blank Sheet</span>',
                        unsafe_allow_html=True)
            fname     = st.text_input("File name", value="Inventory.xlsx", key="bfname")
            fcols_raw = st.text_input("Columns (comma-separated)",
                                      value="Item Name,Qty,Price,Status", key="bcols")
            if st.button("▶  Create & Start", type="primary", use_container_width=True):
                cols = [c.strip() for c in fcols_raw.split(",") if c.strip()]
                fn   = fname.strip() or "Inventory.xlsx"
                if not cols:
                    st.error("Enter at least one column.")
                else:
                    try:
                        ns = SheetStore.create_blank(fn, cols)
                        S().store = ns
                        add_log(f"Created: {ns.file_name}", "ok")
                        go("assistant")   # auto-redirect
                    except Exception as e:
                        st.error(str(e))
            st.markdown('</div>', unsafe_allow_html=True)

    # ── RIGHT: how-it-works + sheet preview ───────────────────────
    with right:
        if loaded():
            s = store()
            # Stats
            st.markdown(f"""
            <div class="stiles">
              <div class="stile">
                <div class="stile-v" style="color:var(--blue);">{len(s.df)}</div>
                <div class="stile-l">Rows</div></div>
              <div class="stile">
                <div class="stile-v" style="color:var(--grn);">{len(s.df.columns)}</div>
                <div class="stile-l">Cols</div></div>
              <div class="stile">
                <div class="stile-v" style="color:var(--amb);">{len(s.sheet_names)}</div>
                <div class="stile-l">Sheets</div></div>
            </div>""", unsafe_allow_html=True)

            # Multi-sheet picker
            if len(s.sheet_names) > 1:
                chosen = st.selectbox("Active sheet", s.sheet_names, key="sp")
                if chosen != s.active_sheet:
                    s.switch_sheet(chosen)

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">📄 Preview</span>', unsafe_allow_html=True)
            st.dataframe(s.df.head(8), use_container_width=True, hide_index=False)
            st.markdown('</div>', unsafe_allow_html=True)
            dl_button()

        else:
            st.markdown("""
            <div class="howbox">
              <div style="font-family:var(--mono);font-size:.56rem;font-weight:700;
                letter-spacing:.14em;text-transform:uppercase;color:var(--blue);
                margin-bottom:10px;">How it works</div>
              <div class="hstep"><div class="hnum">1</div>
                <div class="htxt">Upload <b>.xlsx</b> — auto-redirects to Voice page</div></div>
              <div class="hstep"><div class="hnum">2</div>
                <div class="htxt"><b>Voice:</b> say "Add 50 to Qty where Item is Pen"</div></div>
              <div class="hstep"><div class="hnum">3</div>
                <div class="htxt"><b>Text:</b> type any command directly — no mic needed</div></div>
              <div class="hstep" style="margin-bottom:0;"><div class="hnum">4</div>
                <div class="htxt">Sheet updates instantly · Download .xlsx anytime</div></div>
            </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# PAGE: ASSISTANT  — conversational, preview-before-apply
# Left:  chat history + (pending preview/confirm) + voice/text input
# Right: live sheet preview + download + activity log
# ══════════════════════════════════════════════════════════════════
elif S().page == "assistant":
    if not guard():
        from core import transcribe
        import engine as _eng

        # ── chat-flow helpers (closures over session state) ───────────────
        def _emit_result(pv, mutated=False):
            tables, values, lines = [], [], []
            for s in pv.steps:
                if s.error:
                    lines.append(f"⚠️ {s.summary}: {s.error}")
                elif s.table is not None:
                    tables.append((s.summary, s.table))
                elif s.value is not None:
                    values.append((s.summary, s.value))
                else:
                    lines.append(("✅ " if mutated else "• ") + s.summary)
            text = "\n".join(lines)
            if not text:
                text = "Done — sheet updated." if mutated else "Here you go."
            S().chat.append({"role": "assistant", "kind": "ok",
                             "text": text, "tables": tables, "values": values})

        def _handle(cmd):
            cmd = (cmd or "").strip()
            if not cmd:
                return
            S().chat.append({"role": "user", "text": cmd})
            S().transcript = cmd
            try:
                plan = _eng.interpret(cmd, store().df)
            except Exception as e:
                S().chat.append({"role": "assistant", "kind": "err", "text": f"⚠️ {e}"})
                add_log(str(e), "err")
                return
            if not plan.operations:
                S().chat.append({"role": "assistant", "kind": "info",
                                 "text": plan.summary or
                                 "I couldn't turn that into an action — try rephrasing."})
                return
            if plan.needs_confirm:
                pv = _eng.dry_run(plan, store().df)
                if not pv.ok:
                    S().chat.append({"role": "assistant", "kind": "err",
                                     "text": f"⚠️ {pv.error}"})
                    return
                S().pending = pv
            else:
                pv = _eng.commit(plan, store())
                _emit_result(pv, mutated=False)

        def _confirm():
            pv = S().pending
            S().pending = None
            try:
                done = _eng.commit(pv.plan, store())
            except RuntimeError as e:
                S().chat.append({"role": "assistant", "kind": "err",
                                 "text": f"⚠️ Command ran but **save failed**: {e}"})
                add_log(f"SAVE FAILED: {e}", "err")
                return
            _emit_result(done, mutated=True)
            add_log("Applied: " + (pv.plan.summary or "change"), "ok")

        def _cancel():
            S().pending = None
            S().chat.append({"role": "assistant", "kind": "info",
                             "text": "Cancelled — nothing was changed."})

        left, right = st.columns([1.32, 0.68], gap="large")

        with left:
            # ── header: model + capabilities ──────────────────────────────
            try:
                import llm as _llm
                model = _llm.model_chain()[0] if _llm.is_configured() else "no key — local parser only"
            except Exception:
                model = "—"
            st.markdown(
                f'<div style="display:flex;justify-content:space-between;align-items:center;'
                f'margin-bottom:8px;">'
                f'<span class="clabel" style="margin:0;">💬 Assistant</span>'
                f'<span class="pill pp" style="font-size:.58rem;">🧠 {model}</span></div>',
                unsafe_allow_html=True)

            # ── empty-state hints ─────────────────────────────────────────
            if not S().chat and not S().pending:
                cols = list(store().df.columns)
                ex = []
                if len(cols) >= 2:
                    ex = [f"Increase {cols[1]} by 10% where {cols[0]} is …",
                          f"What is the total of {cols[1]}?",
                          f"Show the top 5 rows by {cols[1]}",
                          f"Add a column Total = {cols[1]} × price",
                          f"Delete rows where {cols[1]} is less than 5"]
                hint = "".join(f'<div style="color:#93c5fd;margin:3px 0;">"{e}"</div>' for e in ex)
                st.markdown(
                    f'<div class="howbox"><div style="font-family:var(--mono);font-size:.56rem;'
                    f'font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--blue);'
                    f'margin-bottom:8px;">Try saying or typing</div>'
                    f'<div style="font-family:var(--mono);font-size:.7rem;line-height:1.7;">{hint}</div>'
                    f'<div style="font-family:var(--mono);font-size:.62rem;color:var(--dim);margin-top:8px;">'
                    f'Ask questions, run filters, sort, add formula columns, or bulk-edit — '
                    f'changes are previewed before they\'re applied.</div></div>',
                    unsafe_allow_html=True)

            # ── chat history ──────────────────────────────────────────────
            for m in S().chat[-24:]:
                if m["role"] == "user":
                    st.markdown(
                        f'<div style="display:flex;justify-content:flex-end;margin:8px 0;">'
                        f'<div style="background:rgba(56,189,248,.14);border:1px solid rgba(56,189,248,.22);'
                        f'border-radius:14px 14px 4px 14px;padding:10px 14px;max-width:85%;'
                        f'font-family:var(--mono);font-size:.82rem;color:#cfe8ff;">{m["text"]}</div></div>',
                        unsafe_allow_html=True)
                else:
                    kind = m.get("kind", "ok")
                    bdr = {"ok": "var(--grn)", "err": "var(--red)",
                           "info": "var(--blue)", "warn": "var(--amb)"}.get(kind, "var(--grn)")
                    body = m["text"].replace("\n", "<br>")
                    st.markdown(
                        f'<div style="display:flex;margin:8px 0;"><div style="background:var(--s1);'
                        f'border:1px solid var(--b1);border-left:3px solid {bdr};'
                        f'border-radius:14px 14px 14px 4px;padding:10px 14px;max-width:92%;'
                        f'font-family:var(--mono);font-size:.82rem;color:var(--txt);">{body}</div></div>',
                        unsafe_allow_html=True)
                    for label, val in m.get("values", []):
                        st.metric(label, val)
                    for title, tdf in m.get("tables", []):
                        st.caption(title)
                        st.dataframe(tdf, use_container_width=True, hide_index=True)

            # ── pending preview / confirm card ────────────────────────────
            if S().pending:
                pv = S().pending
                st.markdown('<div class="card" style="border-color:rgba(251,191,36,.35);">',
                            unsafe_allow_html=True)
                st.markdown('<span class="clabel" style="color:var(--amb);">'
                            '⚠️ Review proposed changes</span>', unsafe_allow_html=True)
                if pv.plan.summary:
                    st.markdown(f'<div class="tx">{pv.plan.summary}</div>', unsafe_allow_html=True)
                _step_rows = []
                for s in pv.steps:
                    aff = (f' <span style="color:var(--amb);">({s.affected} affected)</span>'
                           if s.affected else "")
                    _step_rows.append(
                        f'<div style="font-family:var(--mono);font-size:.74rem;color:var(--txt);'
                        f'padding:4px 0;">• {s.summary}{aff}</div>')
                st.markdown("".join(_step_rows), unsafe_allow_html=True)

                b = pv.before; a = pv.after
                if b is not None and a is not None:
                    st.markdown(
                        f'<div style="font-family:var(--mono);font-size:.66rem;color:var(--dim);'
                        f'margin:8px 0 4px;">Rows {len(b)} → <b style="color:var(--txt);">{len(a)}</b>'
                        f'   ·   Cols {len(b.columns)} → <b style="color:var(--txt);">{len(a.columns)}</b></div>',
                        unsafe_allow_html=True)
                    st.caption("Preview after change")
                    st.dataframe(a.head(8), use_container_width=True, hide_index=True)

                cc1, cc2 = st.columns(2)
                with cc1:
                    st.markdown('<div class="gb">', unsafe_allow_html=True)
                    if st.button("✓ Apply", type="primary", use_container_width=True, key="confirm_btn"):
                        _confirm(); st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                with cc2:
                    st.markdown('<div class="rb">', unsafe_allow_html=True)
                    if st.button("✕ Cancel", use_container_width=True, key="cancel_btn"):
                        _cancel(); st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # ── input area (hidden while a confirmation is pending) ────────
            if not S().pending:
                st.markdown('<hr>', unsafe_allow_html=True)
                m1, m2 = st.columns(2)
                with m1:
                    if st.button("🎙 Voice" + ("  ✓" if S().input_mode == "voice" else ""),
                                 use_container_width=True,
                                 type="primary" if S().input_mode == "voice" else "secondary",
                                 key="am_voice"):
                        S().input_mode = "voice"; S().audio_bytes = None; st.rerun()
                with m2:
                    if st.button("⌨️ Text" + ("  ✓" if S().input_mode == "text" else ""),
                                 use_container_width=True,
                                 type="primary" if S().input_mode == "text" else "secondary",
                                 key="am_text"):
                        S().input_mode = "text"; S().audio_bytes = None; st.rerun()

                if S().input_mode == "voice":
                    audio_val = None
                    if hasattr(st, "audio_input"):
                        audio_val = st.audio_input("Tap the mic, speak your command",
                                                   key="amic", label_visibility="visible")
                    up = st.file_uploader("…or upload an audio clip",
                                          type=["wav", "mp3", "m4a"], key="aud_up")
                    if up is not None:
                        audio_val = up
                    if audio_val is not None:
                        nb = audio_val.read() if hasattr(audio_val, "read") else audio_val
                        if nb and nb != S().audio_bytes:
                            S().audio_bytes = nb
                            with st.spinner("Transcribing…"):
                                try:
                                    t = transcribe(nb)
                                    add_log(f'Heard: "{t[:60]}"', "ok")
                                    _handle(t)
                                    st.rerun()
                                except Exception as e:
                                    st.error(str(e)); add_log(str(e), "err")
                else:
                    typed = st.text_area("Command", key="achat_in", height=72,
                                         placeholder='e.g. "What is the total stock value?"  ·  '
                                                     '"Raise every Writing item\'s price by 15%"',
                                         label_visibility="collapsed")
                    rc, cc = st.columns([3, 1])
                    with rc:
                        if st.button("▶  Send", type="primary", use_container_width=True, key="asend"):
                            if typed.strip():
                                _handle(typed); st.rerun()
                            else:
                                st.warning("Type a command first.")
                    with cc:
                        if st.button("Clear chat", use_container_width=True, key="aclear"):
                            S().chat = []; st.rerun()

        # ── RIGHT COLUMN ──────────────────────────────────────────────────
        with right:
            render_preview(height=300)
            dl_button("⬇️  Download .xlsx")
            if store() and store().undo_depth > 0:
                d = store().undo_depth
                if st.button(f"↩  Undo  ({d} step{'s' if d > 1 else ''})",
                             use_container_width=True, key="aundo"):
                    store().undo(); add_log("Undone ✓", "ok"); st.rerun()
            render_log()



# ══════════════════════════════════════════════════════════════════
# PAGE: DATA
# ══════════════════════════════════════════════════════════════════
elif S().page == "data":
    if not guard():
        s  = store()
        df = s.df
        left, right = st.columns([1.1, 0.9], gap="large")

        with left:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">📋 Sheet Data</span>', unsafe_allow_html=True)
            # ── Metrics row (safe — no undefined variables) ────────
            qty_safe   = pd.to_numeric(df.get("Qty",   pd.Series(dtype=float)), errors="coerce").fillna(0)
            price_safe = pd.to_numeric(df.get("Price", pd.Series(dtype=float)), errors="coerce").fillna(0)
            has_qty   = "Qty"   in df.columns
            has_price = "Price" in df.columns

            if has_qty and has_price:
                m1, m2, m3, m4 = st.columns(4)
                m4.metric("Stock Value", f"₹{(qty_safe * price_safe).sum():,.0f}")
            elif has_qty:
                m1, m2, m3 = st.columns(3)
            else:
                m1, m2, m3 = st.columns(3)

            m1.metric("Rows",    len(df))
            m2.metric("Columns", len(df.columns))
            m3.metric("Total Qty", int(qty_safe.sum()) if has_qty else "—")

            # ── Full table ─────────────────────────────────────────
            st.dataframe(df, use_container_width=True, hide_index=False)
            dl_button()
            st.markdown('</div>', unsafe_allow_html=True)

        with right:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">➕ Add / Edit Row</span>',
                        unsafe_allow_html=True)
            cols_list = df.columns.tolist() if not df.empty \
                        else ["Item Name", "Qty", "Price", "Status"]
            with st.form("row_form"):
                rv = {}
                n  = min(len(cols_list), 3)
                fc = st.columns(n)
                for i, c in enumerate(cols_list):
                    rv[c] = fc[i % n].text_input(c, key=f"rf_{c}")
                if st.form_submit_button("💾  Save Row", type="primary"):
                    first = cols_list[0]
                    v0    = rv.get(first, "").strip()
                    if not v0:
                        st.warning("First column value is required.")
                    else:
                        lu = next((c for c in cols_list
                                   if c.strip().lower() == "last updated"), None)
                        if lu:
                            rv[lu] = datetime.now().strftime("%Y-%m-%d %H:%M")
                        df2  = s.df
                        mask = (df2[first].astype(str).str.strip().str.lower() == v0.lower()
                                if first in df2.columns
                                else pd.Series([False] * len(df2)))
                        try:
                            if mask.any():
                                s.update_row(int(df2[mask].index[0]),
                                             {k: v for k, v in rv.items() if v not in ("", None)})
                                add_log(f"Updated: {v0}", "ok")
                                st.success("✅ Row updated and saved to disk!")
                            else:
                                s.add_row(rv)
                                add_log(f"Added: {v0}", "ok")
                                st.success("✅ Row added and saved to disk!")
                            st.rerun()
                        except RuntimeError as save_err:
                            add_log(f"SAVE FAILED: {save_err}", "err")
                            st.error(
                                f"⚠️ Row was added to memory but **save failed**: {save_err}\n\n"
                                "Close the file in Excel if it is open, then try again."
                            )
            st.markdown('</div>', unsafe_allow_html=True)
            render_log()


# ══════════════════════════════════════════════════════════════════
# PAGE: DELETE
# ══════════════════════════════════════════════════════════════════
elif S().page == "delete":
    if not guard():
        s  = store()
        df = s.df
        ca, cb = st.columns(2, gap="large")

        with ca:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">🗑 Delete Rows</span>',
                        unsafe_allow_html=True)
            if df.empty:
                st.markdown('<span class="pill pa">No data</span>',
                            unsafe_allow_html=True)
            else:
                first  = df.columns[0]
                labels = df[first].fillna("(blank)").astype(str).tolist()
                to_del = st.multiselect("Select rows to delete", labels)
                st.markdown('<div class="rb">', unsafe_allow_html=True)
                if st.button("🗑  Delete Selected", use_container_width=True) and to_del:
                    idxs = [i for i, v in enumerate(labels) if v in to_del]
                    s.delete_rows(idxs)
                    add_log(f"Deleted {len(idxs)} row(s)", "ok")
                    st.success(f"Deleted {len(idxs)} row(s).")
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with cb:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">🧹 Clear All Rows</span>',
                        unsafe_allow_html=True)
            st.markdown(
                '<p style="font-size:.75rem;color:var(--dim);margin-bottom:10px;">'
                'Removes all rows. Column headers are preserved.</p>',
                unsafe_allow_html=True)
            confirm = st.checkbox("I understand — remove all rows")
            st.markdown('<div class="rb">', unsafe_allow_html=True)
            if st.button("⚠️  Clear Sheet", use_container_width=True) and confirm:
                s.clear_rows()
                add_log("Sheet cleared", "err")
                st.success("Cleared.")
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        if s.undo_depth > 0:
            if st.button(f"↩  Undo  ({s.undo_depth} step(s))", use_container_width=True):
                s.undo()
                add_log("Undone ✓", "ok")
                st.rerun()


# ══════════════════════════════════════════════════════════════════
# PAGE: COLUMNS
# ══════════════════════════════════════════════════════════════════
elif S().page == "columns":
    if not guard():
        s  = store()
        df = s.df
        ca, cb = st.columns(2, gap="large")

        with ca:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">📐 Current Columns</span>',
                        unsafe_allow_html=True)
            for i, c in enumerate(df.columns.tolist()):
                st.markdown(
                    f'<div class="col-row">'
                    f'<span class="col-idx">{i+1:02d}</span>{c}</div>',
                    unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">➕ Add Column</span>',
                        unsafe_allow_html=True)
            nc = st.text_input("Name", key="add_col",
                               label_visibility="collapsed",
                               placeholder="e.g. Supplier")
            if st.button("➕  Add", use_container_width=True) and nc.strip():
                try:
                    s.add_column(nc.strip())
                    add_log(f"Added col: {nc}", "ok")
                    st.success("Added!")
                    st.rerun()
                except ValueError as e:
                    st.warning(str(e))
            st.markdown('</div>', unsafe_allow_html=True)

        with cb:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">✏️ Rename Column</span>',
                        unsafe_allow_html=True)
            oc = st.selectbox("Column", ["—"] + df.columns.tolist(), key="ren_sel")
            rn = st.text_input("New name", key="ren_new")
            if st.button("✏️  Rename", use_container_width=True) \
               and oc != "—" and rn.strip():
                try:
                    s.rename_column(oc, rn.strip())
                    add_log(f"Renamed '{oc}' → '{rn}'", "ok")
                    st.success("Renamed!")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">🗑 Delete Column</span>',
                        unsafe_allow_html=True)
            dc = st.selectbox("Column to delete", ["—"] + df.columns.tolist(),
                              key="del_col")
            st.markdown('<div class="rb">', unsafe_allow_html=True)
            if st.button("🗑  Delete", use_container_width=True) and dc != "—":
                try:
                    s.delete_column(dc)
                    add_log(f"Deleted col: {dc}", "ok")
                    st.success("Deleted!")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))
            st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════════════
elif S().page == "dashboard":
    if not guard():
        s  = store()
        df = s.df
        left, right = st.columns([1, 1], gap="large")

        with left:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">📊 Overview</span>', unsafe_allow_html=True)
            if df.empty:
                st.markdown('<span class="pill pa">No data yet</span>',
                            unsafe_allow_html=True)
            else:
                qty_s   = pd.to_numeric(df.get("Qty",   pd.Series(dtype=float)), errors="coerce").fillna(0)
                price_s = pd.to_numeric(df.get("Price", pd.Series(dtype=float)), errors="coerce").fillna(0)
                has_q = "Qty"   in df.columns
                has_p = "Price" in df.columns

                if has_q and has_p:
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Rows",        len(df))
                    m2.metric("Columns",     len(df.columns))
                    m3.metric("Total Qty",   int(qty_s.sum()))
                    m4.metric("Stock Value", f"₹{(qty_s * price_s).sum():,.0f}")
                elif has_q:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Rows",      len(df))
                    m2.metric("Columns",   len(df.columns))
                    m3.metric("Total Qty", int(qty_s.sum()))
                else:
                    m1, m2 = st.columns(2)
                    m1.metric("Rows",    len(df))
                    m2.metric("Columns", len(df.columns))
            st.markdown('</div>', unsafe_allow_html=True)

            # ── Save & Export card ──────────────────────────────────
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">💾 Save & Export</span>', unsafe_allow_html=True)
            s2 = store()
            if s2 and s2.file_bytes:
                import subprocess, sys as _sys

                save_path = s2.file_path or ""

                # File path display
                st.markdown(
                    f'<div style="font-family:var(--mono);font-size:.62rem;color:var(--dim);'
                    f'margin-bottom:12px;line-height:1.6;">'
                    f'📁 <b style="color:var(--txt);">{s2.file_name}</b>'
                    f' · {len(s2.file_bytes)//1024} KB<br>'
                    f'<span style="color:var(--grn);word-break:break-all;">{save_path}</span>'
                    f'</div>',
                    unsafe_allow_html=True)

                # ── OPEN IN EXCEL button ───────────────────────────
                # Uses subprocess to launch Excel with the exact saved file path.
                # This works because Streamlit runs on the same machine as the user.
                # No download, no browser URI tricks — just opens the file directly.
                if st.button("📊  Open in Excel", type="primary",
                             use_container_width=True, key="open_excel_btn"):
                    if not save_path or not os.path.exists(save_path):
                        st.error("File path not found. Save the file first.")
                    else:
                        try:
                            if _sys.platform == "win32":
                                os.startfile(save_path)          # Windows — opens in default app (Excel)
                            elif _sys.platform == "darwin":
                                subprocess.Popen(["open", save_path])   # macOS
                            else:
                                subprocess.Popen(["xdg-open", save_path])  # Linux
                            st.success(f"✅ Opening {s2.file_name} in Excel…")
                            add_log(f"Opened in Excel: {save_path}", "ok")
                        except Exception as ex:
                            st.error(f"Could not open Excel: {ex}")
                            add_log(str(ex), "err")

                st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

                # ── Download button (backup only) ──────────────────
                st.download_button(
                    label="⬇️  Download .xlsx (backup)",
                    data=s2.file_bytes,
                    file_name=s2.file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dash_dl_only",
                )

                st.markdown(
                    '<div style="font-family:var(--mono);font-size:.59rem;color:var(--dim);' +
                    'margin-top:8px;line-height:1.7;">' +
                    '· <b style="color:var(--blue);">Open in Excel</b> launches the saved file directly — no download<br>' +
                    '· <b style="color:var(--sub);">Download</b> saves a backup copy to your Downloads folder' +
                    '</div>',
                    unsafe_allow_html=True)

            else:
                st.markdown('<span class="pill pa">No file loaded</span>',
                            unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with right:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<span class="clabel">📈 Charts</span>', unsafe_allow_html=True)
            if not df.empty:
                num_cols = df.select_dtypes("number").columns.tolist()
                first    = df.columns[0]
                for nc in num_cols[:3]:
                    cdf = df[[first, nc]].dropna().set_index(first)
                    if not cdf.empty:
                        st.markdown(
                            f'<div style="font-family:var(--mono);font-size:.6rem;'
                            f'color:var(--dim);text-transform:uppercase;letter-spacing:.1em;'
                            f'margin:8px 0 2px;">{nc}</div>',
                            unsafe_allow_html=True)
                        st.bar_chart(cdf[[nc]], height=150)
            else:
                st.markdown('<span class="pill pa">No data yet</span>',
                            unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)