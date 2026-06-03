"""
SheetMind Excel Add-in — xlwings Python module.
Called from Excel ribbon buttons via VBA RunPython.
"""
from __future__ import annotations
import os
import sys
import pandas as pd
from dotenv import load_dotenv

# Ensure core.py is importable (xlwings sets cwd to the xlam directory)
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

load_dotenv(os.path.join(_HERE, '.env'))

import xlwings as xw

# In-process undo history (lives as long as Excel session)
_undo_stack: list[pd.DataFrame] = []
_UNDO_LIMIT = 25


# ══════════════════════════════════════════════════════════════════
# XlwingsStore — drop-in replacement for SheetStore that reads/writes
# the active Excel sheet directly instead of a file on disk.
# Implements the same interface that core.apply_command() expects.
# ══════════════════════════════════════════════════════════════════
class XlwingsStore:
    def __init__(self, ws: xw.Sheet, df: pd.DataFrame):
        self._ws = ws
        self._df = df.copy()

    @property
    def df(self) -> pd.DataFrame:
        return self._df

    def save(self) -> None:
        self._ws.clear_contents()
        if not self._df.empty:
            rows = [self._df.columns.tolist()] + self._df.values.tolist()
        else:
            rows = [self._df.columns.tolist()]
        self._ws['A1'].value = rows

    def _push_history(self) -> None:
        _undo_stack.append(self._df.copy(deep=True))
        if len(_undo_stack) > _UNDO_LIMIT:
            _undo_stack.pop(0)

    def undo(self) -> bool:
        if not _undo_stack:
            return False
        self._df = _undo_stack.pop()
        self.save()
        return True

    def add_row(self, data: dict) -> None:
        self._push_history()
        row = {c: data.get(c, '') for c in self._df.columns}
        self._df = pd.concat([self._df, pd.DataFrame([row])], ignore_index=True)
        self.save()

    def update_row(self, idx: int, data: dict) -> None:
        self._push_history()
        df = self._df.copy()
        for col, val in data.items():
            if col in df.columns:
                df.at[idx, col] = val
        self._df = df
        self.save()

    def delete_rows(self, indices: list) -> None:
        self._push_history()
        self._df = self._df.drop(index=indices).reset_index(drop=True)
        self.save()

    def delete_rows_matching(self, col: str, value: str) -> int:
        if col not in self._df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(self._df.columns)}")
        self._push_history()
        mask = self._df[col].astype(str).str.strip().str.lower() == str(value).strip().lower()
        count = int(mask.sum())
        self._df = self._df[~mask].reset_index(drop=True)
        self.save()
        return count

    def update_cells_matching(self, match_col: str, match_val: str,
                               target_col: str, new_val) -> int:
        for c in (match_col, target_col):
            if c not in self._df.columns:
                raise ValueError(f"Column '{c}' not found. Available: {list(self._df.columns)}")
        self._push_history()
        df = self._df.copy()
        mask = (df[match_col].astype(str).str.strip().str.lower()
                == str(match_val).strip().lower())
        if isinstance(new_val, str) and new_val.startswith(('+', '-')):
            try:
                delta = float(new_val)
                df.loc[mask, target_col] = (
                    pd.to_numeric(df.loc[mask, target_col], errors='coerce').fillna(0) + delta
                )
            except Exception:
                df.loc[mask, target_col] = new_val
        else:
            df.loc[mask, target_col] = new_val
        self._df = df
        self.save()
        return int(mask.sum())

    def clear_rows(self) -> None:
        self._push_history()
        self._df = pd.DataFrame(columns=self._df.columns)
        self.save()


# ══════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════
def _read_active_sheet():
    """Return (wb, ws, df) from the caller's active sheet."""
    wb = xw.Book.caller()
    ws = wb.sheets.active
    used = ws.used_range
    if used is None:
        return wb, ws, pd.DataFrame()

    data = used.value
    if not data:
        return wb, ws, pd.DataFrame()

    # Normalize single-row edge case
    if not isinstance(data[0], list):
        data = [data]

    headers = [str(h).strip() if h is not None else f'Col{i}'
               for i, h in enumerate(data[0])]
    rows = [r for r in data[1:]
            if r is not None and isinstance(r, list) and any(v is not None for v in r)]

    df = pd.DataFrame(rows, columns=headers) if rows else pd.DataFrame(columns=headers)

    for col in df.columns:
        converted = pd.to_numeric(df[col], errors='coerce')
        if converted.notna().sum() > 0 and converted.notna().mean() > 0.5:
            df[col] = converted

    return wb, ws, df


def _msgbox(app, msg: str, title: str = 'SheetMind', icon: int = 64) -> None:
    """icon: 64=info, 48=warning, 16=error"""
    try:
        app.api.MsgBox(msg, icon, title)
    except Exception:
        pass


def _inputbox(app, prompt: str, title: str = 'SheetMind',
              default: str = '') -> str | None:
    try:
        result = app.api.InputBox(prompt, title, default, Type=2)
        if result is False or result == '':
            return None
        return str(result)
    except Exception:
        return None


def _confirm_box(app, msg: str, title: str = 'SheetMind — Confirm') -> bool:
    """Yes/No dialog. vbYesNo(4)+vbQuestion(32)=36; vbYes=6."""
    try:
        return int(app.api.MsgBox(msg, 36, title)) == 6
    except Exception:
        return False


def _format_results(steps) -> str:
    lines = []
    for s in steps:
        if s.error:
            lines.append(f'⚠️ {s.summary}: {s.error}')
        elif s.value is not None:
            lines.append(f'{s.summary}\n   = {s.value}')
        elif s.table is not None:
            lines.append(s.summary + '\n' + s.table.to_string(index=False))
        else:
            lines.append('✅ ' + s.summary)
    return '\n'.join(lines) or 'Done.'


def _preview_text(plan, pv) -> str:
    lines = [plan.summary or 'Proposed changes:', '']
    for s in pv.steps:
        suffix = f'  ({s.affected} affected)' if s.affected else ''
        lines.append(f'• {s.summary}{suffix}')
    if pv.before is not None and pv.after is not None:
        lines += ['', f'Rows {len(pv.before)} → {len(pv.after)}   '
                      f'Cols {len(pv.before.columns)} → {len(pv.after.columns)}']
    lines += ['', 'Apply these changes?']
    return '\n'.join(lines)


# ══════════════════════════════════════════════════════════════════
# Premium native dialogs — same look as the SheetMind web pane:
# emerald palette (auto dark/light per OS), spark logo, gradient header,
# sleek buttons, styled preview table. No browser/server needed.
# ══════════════════════════════════════════════════════════════════
def _os_is_dark() -> bool:
    try:
        if sys.platform == 'darwin':
            import subprocess
            r = subprocess.run(['defaults', 'read', '-g', 'AppleInterfaceStyle'],
                               capture_output=True, text=True)
            return 'dark' in (r.stdout or '').lower()
        if sys.platform == 'win32':
            import winreg
            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
            return winreg.QueryValueEx(k, 'AppsUseLightTheme')[0] == 0
    except Exception:
        pass
    return True


def _palette() -> dict:
    if _os_is_dark():
        return dict(BG='#0f1115', CARD='#15181e', CARD2='#1b1f27', TXT='#eef1f4',
                    SUB='#9aa3ad', DIM='#646c77', HAIR='#262b33', ZEBRA='#181b21',
                    ACCENT='#34d399', ACCENT2='#10b981', INK='#052e22',
                    RED='#fb7185', AMB='#fbbf24', HEAD='#ffffff', SUBHEAD='#dffaf0')
    return dict(BG='#f4f6f4', CARD='#ffffff', CARD2='#f3f5f4', TXT='#10221c',
                SUB='#566169', DIM='#8b949c', HAIR='#e1e6e2', ZEBRA='#f7f9f8',
                ACCENT='#059669', ACCENT2='#047857', INK='#ffffff',
                RED='#e11d48', AMB='#b45309', HEAD='#ffffff', SUBHEAD='#eafff5')


def _fonts(root):
    import tkinter.font as tkfont
    fams = set(tkfont.families(root))
    def pick(cands):
        for c in cands:
            if c in fams:
                return c
        return cands[-1]
    disp = pick(['SF Pro Display', 'Helvetica Neue', 'Segoe UI Semibold', 'Segoe UI', 'Arial'])
    body = pick(['SF Pro Text', 'Helvetica Neue', 'Segoe UI', 'Arial'])
    mono = pick(['JetBrains Mono', 'SF Mono', 'Menlo', 'Consolas', 'Courier New'])
    return disp, body, mono


def _hex(c):
    c = c.lstrip('#')
    return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)


def _center(root, w, h):
    root.update_idletasks()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 3
    root.geometry(f'{w}x{h}+{x}+{y}')


def _make_window(title):
    import tkinter as tk
    root = tk.Tk()
    root.title(title)
    P = _palette()
    root.configure(bg=P['BG'])
    root.attributes('-topmost', True)
    root.resizable(False, False)
    return root, P, _fonts(root)


def _gradient_header(root, P, disp, mono, title, subtitle, w):
    """Horizontal emerald→teal gradient banner with the spark logo."""
    import tkinter as tk
    h = 64
    cv = tk.Canvas(root, width=w, height=h, highlightthickness=0, bd=0)
    cv.pack(fill='x')
    (r1, g1, b1), (r2, g2, b2) = _hex(P['ACCENT']), _hex(P['ACCENT2'])
    for x in range(w):
        t = x / max(1, w - 1)
        cv.create_line(x, 0, x, h, fill=f'#{int(r1+(r2-r1)*t):02x}'
                       f'{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}')
    tx = 22
    try:
        path = os.path.join(_HERE, 'addin_web', 'assets', 'icon-128.png')
        img = tk.PhotoImage(master=root, file=path)
        f = max(1, img.width() // 34)
        img = img.subsample(f, f)
        cv.create_image(20, h // 2, image=img, anchor='w')
        root._sm_logo = img  # keep a ref so it isn't garbage-collected
        tx = 64
    except Exception:
        pass
    cv.create_text(tx, h // 2 - 9, text=title, anchor='w', fill=P['HEAD'], font=(disp, 15, 'bold'))
    cv.create_text(tx, h // 2 + 11, text=subtitle, anchor='w', fill=P['SUBHEAD'], font=(mono, 9))
    return cv


def _btn(parent, text, cmd, P, body, kind='primary'):
    import tkinter as tk
    base = dict(relief='flat', cursor='hand2', bd=0, highlightthickness=0,
                font=(body, 11, 'bold'), padx=20, pady=9)
    if kind == 'primary':
        b = tk.Button(parent, text=text, command=cmd, bg=P['ACCENT'], fg=P['INK'],
                      activebackground=P['ACCENT2'], activeforeground=P['INK'], **base)
    elif kind == 'danger':
        b = tk.Button(parent, text=text, command=cmd, bg=P['CARD2'], fg=P['RED'],
                      activebackground=P['CARD'], activeforeground=P['RED'], **base)
    else:
        b = tk.Button(parent, text=text, command=cmd, bg=P['CARD2'], fg=P['TXT'],
                      activebackground=P['CARD'], activeforeground=P['TXT'], **base)
    return b


def _card(parent, P):
    """A sleek surface 'card' (1px hairline border via an outer frame)."""
    import tkinter as tk
    outer = tk.Frame(parent, bg=P['HAIR'])
    inner = tk.Frame(outer, bg=P['CARD'])
    inner.pack(fill='both', expand=True, padx=1, pady=1)
    return outer, inner


def _make_table(parent, P, df, max_rows=10):
    from tkinter import ttk
    a = df.head(max_rows)
    cols = [str(c) for c in a.columns]
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure('SM.Treeview', background=P['CARD'], fieldbackground=P['CARD'],
                    foreground=P['TXT'], rowheight=26, borderwidth=0, font=('TkDefaultFont', 10))
    style.configure('SM.Treeview.Heading', background=P['CARD2'], foreground=P['SUB'],
                    relief='flat', borderwidth=0, font=('TkDefaultFont', 9, 'bold'))
    style.map('SM.Treeview', background=[('selected', P['ACCENT'])],
              foreground=[('selected', P['INK'])])
    tv = ttk.Treeview(parent, columns=cols, show='headings',
                      height=min(max_rows, max(1, len(a))), style='SM.Treeview')
    width = max(80, min(160, 580 // max(1, len(cols))))
    for c in cols:
        tv.heading(c, text=str(c).upper())
        tv.column(c, width=width, anchor='w')
    tv.tag_configure('odd', background=P['ZEBRA'])
    for i, (_, r) in enumerate(a.iterrows()):
        tv.insert('', 'end', values=[r[c] for c in a.columns], tags=('odd',) if i % 2 else ())
    return tv


def _styled_input(prompt_title: str, examples: str = '') -> str | None:
    """Premium command entry. Returns the typed command or None if cancelled."""
    import tkinter as tk
    W = 560
    root, P, (disp, body, mono) = _make_window('SheetMind')
    out = {'cmd': None}

    _gradient_header(root, P, disp, mono, 'SheetMind', 'copilot for excel', W)

    wrap = tk.Frame(root, bg=P['BG'])
    wrap.pack(fill='both', expand=True, padx=22, pady=18)
    tk.Label(wrap, text=prompt_title, bg=P['BG'], fg=P['TXT'],
             font=(disp, 14, 'bold')).pack(anchor='w')
    if examples:
        tk.Label(wrap, text=examples, bg=P['BG'], fg=P['SUB'], font=(mono, 9),
                 justify='left').pack(anchor='w', pady=(6, 12))

    outer, inner = _card(wrap, P)
    outer.pack(fill='x')
    entry = tk.Text(inner, height=3, bg=P['CARD'], fg=P['TXT'], insertbackground=P['ACCENT'],
                    relief='flat', font=(mono, 11), wrap='word', padx=12, pady=10, bd=0,
                    highlightthickness=0)
    entry.pack(fill='x')
    entry.focus()

    def run(_=None):
        out['cmd'] = entry.get('1.0', 'end').strip()
        root.destroy()

    bf = tk.Frame(wrap, bg=P['BG'])
    bf.pack(fill='x', pady=(16, 0))
    _btn(bf, 'Run  ▶', run, P, body, 'primary').pack(side='left')
    _btn(bf, 'Cancel', root.destroy, P, body, 'ghost').pack(side='left', padx=8)
    root.bind('<Control-Return>', run)
    _center(root, W, 290)
    root.mainloop()
    return out['cmd'] or None


def _styled_preview(plan, pv) -> bool:
    """Visual review of pending changes with a preview table. Returns True to apply."""
    import tkinter as tk
    W = 620
    root, P, (disp, body, mono) = _make_window('SheetMind — Review')
    out = {'ok': False}

    _gradient_header(root, P, disp, mono, 'Review changes', 'preview before applying', W)

    wrap = tk.Frame(root, bg=P['BG'])
    wrap.pack(fill='both', expand=True, padx=22, pady=16)
    if plan.summary:
        tk.Label(wrap, text=plan.summary, bg=P['BG'], fg=P['TXT'], font=(body, 11),
                 wraplength=W - 60, justify='left').pack(anchor='w', pady=(0, 10))

    for s in pv.steps:
        suffix = f'   ({s.affected} affected)' if s.affected else ''
        row = tk.Frame(wrap, bg=P['BG'])
        row.pack(fill='x', pady=1)
        tk.Label(row, text='•', bg=P['BG'], fg=P['ACCENT'], font=(body, 11)).pack(side='left')
        tk.Label(row, text=f' {s.summary}{suffix}', bg=P['BG'], fg=P['SUB'],
                 font=(mono, 9), justify='left', wraplength=W - 80).pack(side='left')

    if pv.before is not None and pv.after is not None:
        tk.Label(wrap, text=f'Rows {len(pv.before)} → {len(pv.after)}      '
                            f'Cols {len(pv.before.columns)} → {len(pv.after.columns)}',
                 bg=P['BG'], fg=P['ACCENT'], font=(mono, 9)).pack(anchor='w', pady=(10, 4))
        outer, inner = _card(wrap, P)
        outer.pack(fill='x')
        _make_table(inner, P, pv.after).pack(fill='x')

    def apply():
        out['ok'] = True
        root.destroy()

    bf = tk.Frame(wrap, bg=P['BG'])
    bf.pack(fill='x', pady=(16, 0))
    _btn(bf, '✓  Apply', apply, P, body, 'primary').pack(side='left', fill='x', expand=True)
    _btn(bf, '✕  Cancel', root.destroy, P, body, 'danger').pack(side='left', padx=(10, 0), fill='x', expand=True)
    _center(root, W, 520)
    root.mainloop()
    return out['ok']


def _styled_result(steps, title: str = 'Result') -> None:
    """Show command results (answers + tables) in a styled window."""
    import tkinter as tk
    W = 600
    applied = title.lower().startswith('applied')
    root, P, (disp, body, mono) = _make_window(f'SheetMind — {title}')

    _gradient_header(root, P, disp, mono,
                     'Done' if applied else 'Result',
                     'changes applied to your sheet' if applied else 'here you go', W)

    wrap = tk.Frame(root, bg=P['BG'])
    wrap.pack(fill='both', expand=True, padx=22, pady=16)

    for s in steps:
        if s.error:
            tk.Label(wrap, text=f'⚠  {s.summary}: {s.error}', bg=P['BG'], fg=P['RED'],
                     font=(mono, 9), wraplength=W - 60, justify='left').pack(anchor='w', pady=4)
        elif s.value is not None:
            outer, inner = _card(wrap, P)
            outer.pack(fill='x', pady=4)
            pad = tk.Frame(inner, bg=P['CARD'])
            pad.pack(fill='x', padx=14, pady=12)
            tk.Label(pad, text=s.summary, bg=P['CARD'], fg=P['SUB'], font=(mono, 9)).pack(anchor='w')
            tk.Label(pad, text=str(s.value), bg=P['CARD'], fg=P['ACCENT'],
                     font=(disp, 22, 'bold')).pack(anchor='w')
        elif s.table is not None:
            tk.Label(wrap, text=s.summary, bg=P['BG'], fg=P['SUB'],
                     font=(mono, 9)).pack(anchor='w', pady=(8, 4))
            outer, inner = _card(wrap, P)
            outer.pack(fill='x')
            _make_table(inner, P, s.table, max_rows=12).pack(fill='x')
        else:
            tk.Label(wrap, text='✓  ' + s.summary, bg=P['BG'], fg=P['TXT'],
                     font=(body, 11), wraplength=W - 60, justify='left').pack(anchor='w', pady=3)

    bf = tk.Frame(wrap, bg=P['BG'])
    bf.pack(fill='x', pady=(16, 0))
    _btn(bf, 'Close', root.destroy, P, body, 'primary').pack()
    _center(root, W, 460)
    root.mainloop()


def _process_command(app, ws, df, cmd: str) -> None:
    """Shared pipeline for ribbon text/voice commands: interpret → visual review →
    confirm (for mutations) → commit, all via the SheetMind engine."""
    import engine
    store = XlwingsStore(ws, df)
    try:
        plan = engine.interpret(cmd, store.df)
    except Exception as e:
        _msgbox(app, f'Could not understand the command:\n{e}', icon=48)
        return
    if not plan.operations:
        _msgbox(app, plan.summary or 'No action recognized — try rephrasing.', icon=48)
        return

    pv = engine.dry_run(plan, store.df)
    if not pv.ok:
        _msgbox(app, f'Cannot apply:\n{pv.error}', icon=48)
        return

    if plan.needs_confirm:
        if not _styled_preview(plan, pv):       # visual review + confirm
            return
        done = engine.commit(plan, store)
        _styled_result(done.steps, title='Applied')
    else:
        done = engine.commit(plan, store)       # read-only — no save
        _styled_result(done.steps, title='Result')


# ══════════════════════════════════════════════════════════════════
# Public ribbon functions — called via VBA  RunPython / Application.Run
# ══════════════════════════════════════════════════════════════════

def run_text_command():
    """Ribbon → Run Command: InputBox → parse → apply → result."""
    try:
        wb, ws, df = _read_active_sheet()
        app = wb.app

        if df.empty and not list(df.columns):
            _msgbox(app,
                    'No data found in the active sheet.\n'
                    'Make sure row 1 contains column headers.',
                    icon=48)
            return

        cols = list(df.columns)
        hint = ''
        if len(cols) >= 2:
            hint = (f'Examples for this sheet:\n'
                    f'  • Increase {cols[1]} by 10% where {cols[0]} is Example\n'
                    f'  • What is the total {cols[1]}?\n'
                    f'  • Show the top 5 rows by {cols[1]}\n'
                    f'  • Delete rows where {cols[1]} is less than 5\n'
                    f'  (Ctrl+Enter to run)')

        cmd = _styled_input('What should I do?', hint)
        if not cmd:
            return

        _process_command(app, ws, df, cmd)

    except Exception as e:
        try:
            xw.Book.caller().app.api.MsgBox(f'Error:\n{e}', 16, 'SheetMind')
        except Exception:
            pass


def run_voice_command():
    """Ribbon → Voice: record from mic → transcribe via Whisper → confirm/edit → apply."""
    try:
        wb, ws, df = _read_active_sheet()
        app = wb.app

        if df.empty and not list(df.columns):
            _msgbox(app, 'No data found in the active sheet.', icon=48)
            return

        # Ensure sounddevice is available
        try:
            import sounddevice as sd
            import numpy as np
        except ImportError:
            import subprocess
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'sounddevice', 'numpy'],
                          check=True, capture_output=True)
            import sounddevice as sd
            import numpy as np

        import tkinter as tk
        import threading
        import wave
        import tempfile

        SAMPLE_RATE = 16000  # Whisper works best at 16 kHz

        # ── Recording window ────────────────────────────────────────────────
        W = 420
        root = tk.Tk()
        root.title('SheetMind — Voice')
        P = _palette(); disp, body, mono = _fonts(root)
        root.configure(bg=P['BG'])
        root.resizable(False, False)
        root.attributes('-topmost', True)

        recording_chunks: list = []
        is_recording = threading.Event()
        result = {'transcript': None}

        _gradient_header(root, P, disp, mono, 'Voice command', 'speak; Whisper transcribes it', W)
        wrap = tk.Frame(root, bg=P['BG'])
        wrap.pack(fill='both', expand=True, padx=22, pady=18)

        status_var = tk.StringVar(value='Press record and speak your command')
        status_lbl = tk.Label(wrap, textvariable=status_var, bg=P['BG'], fg=P['SUB'],
                              font=(body, 10), wraplength=W - 60)
        status_lbl.pack(pady=(0, 16))

        btn_frame = tk.Frame(wrap, bg=P['BG'])
        btn_frame.pack()
        btn_text = tk.StringVar(value='●   Start recording')
        rec_btn = tk.Button(btn_frame, textvariable=btn_text, font=(body, 11, 'bold'),
                            padx=22, pady=10, bd=0, relief='flat', cursor='hand2',
                            bg=P['ACCENT'], fg=P['INK'], activebackground=P['ACCENT2'],
                            activeforeground=P['INK'])
        rec_btn.pack(side='left', padx=6)
        _btn(btn_frame, 'Cancel', root.destroy, P, body, 'ghost').pack(side='left', padx=6)

        def _record_worker():
            with sd.InputStream(samplerate=SAMPLE_RATE, channels=1, dtype='int16') as stream:
                while is_recording.is_set():
                    chunk, _ = stream.read(1024)
                    recording_chunks.append(chunk.copy())

        def _do_transcribe():
            try:
                audio_data = np.concatenate(recording_chunks, axis=0)
                tmp = tempfile.NamedTemporaryFile(suffix='.wav', delete=False)
                tmp.close()
                with wave.open(tmp.name, 'wb') as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(SAMPLE_RATE)
                    wf.writeframes(audio_data.tobytes())
                with open(tmp.name, 'rb') as f:
                    audio_bytes = f.read()
                import os; os.unlink(tmp.name)

                from core import transcribe
                transcript = transcribe(audio_bytes)
                result['transcript'] = transcript
                root.after(0, root.destroy)
            except Exception as e:
                root.after(0, lambda: (
                    status_var.set(f'Error: {e}'),
                    status_lbl.config(fg=P['RED']),
                    rec_btn.config(state='normal', bg=P['ACCENT'], text='●   Start recording'),
                    btn_text.set('●   Start recording'),
                ))

        def toggle():
            if not is_recording.is_set():
                # START
                recording_chunks.clear()
                is_recording.set()
                threading.Thread(target=_record_worker, daemon=True).start()
                btn_text.set('■   Stop & transcribe')
                rec_btn.config(bg=P['RED'], activebackground=P['RED'])
                status_var.set('Recording…  speak your command now')
                status_lbl.config(fg=P['RED'])
            else:
                # STOP
                is_recording.clear()
                rec_btn.config(state='disabled')
                btn_text.set('Transcribing…')
                status_var.set('Sending to Whisper — please wait…')
                status_lbl.config(fg=P['ACCENT'])
                threading.Thread(target=_do_transcribe, daemon=True).start()

        rec_btn.config(command=toggle)
        _center(root, W, 250)
        root.mainloop()

        if result['transcript'] is None:
            return

        transcript = result['transcript']

        # ── Edit / confirm window ──────────────────────────────────────────
        W2 = 480
        root2 = tk.Tk()
        root2.title('SheetMind — Confirm')
        P2 = _palette(); d2, b2, m2 = _fonts(root2)
        root2.configure(bg=P2['BG'])
        root2.resizable(False, False)
        root2.attributes('-topmost', True)

        final = {'cmd': None}

        _gradient_header(root2, P2, d2, m2, 'Confirm command', 'edit if needed, then run', W2)
        wrap2 = tk.Frame(root2, bg=P2['BG'])
        wrap2.pack(fill='both', expand=True, padx=22, pady=18)
        tk.Label(wrap2, text='Whisper heard:', bg=P2['BG'], fg=P2['SUB'],
                 font=(b2, 10)).pack(anchor='w', pady=(0, 6))

        outer, inner = _card(wrap2, P2)
        outer.pack(fill='x')
        entry = tk.Entry(inner, font=(m2, 11), bg=P2['CARD'], fg=P2['TXT'],
                         insertbackground=P2['ACCENT'], relief='flat', bd=0,
                         highlightthickness=0)
        entry.insert(0, transcript)
        entry.select_range(0, 'end')
        entry.pack(fill='x', ipady=9, padx=10)
        entry.focus()

        def confirm():
            final['cmd'] = entry.get().strip()
            root2.destroy()

        f2 = tk.Frame(wrap2, bg=P2['BG'])
        f2.pack(fill='x', pady=(16, 0))
        _btn(f2, 'Run  ▶', confirm, P2, b2, 'primary').pack(side='left')
        _btn(f2, 'Cancel', root2.destroy, P2, b2, 'ghost').pack(side='left', padx=8)
        root2.bind('<Return>', lambda _: confirm())
        _center(root2, W2, 240)
        root2.mainloop()

        if not final['cmd']:
            return

        _process_command(app, ws, df, final['cmd'])

    except Exception as e:
        try:
            xw.Book.caller().app.api.MsgBox(f'Error:\n{e}', 16, 'SheetMind')
        except Exception:
            pass


def undo_last():
    """Ribbon → Undo: revert last SheetMind change."""
    try:
        wb, ws, df = _read_active_sheet()
        app = wb.app

        if not _undo_stack:
            _msgbox(app,
                    'Nothing to undo — no SheetMind changes recorded this session.',
                    icon=64)
            return

        store = XlwingsStore(ws, df)
        store.undo()
        _msgbox(app,
                f'✅ Undone! ({len(_undo_stack)} step(s) still available)',
                icon=64)

    except Exception as e:
        try:
            xw.Book.caller().app.api.MsgBox(f'Undo error:\n{e}', 16, 'SheetMind')
        except Exception:
            pass


def show_info():
    """Ribbon → Sheet Info: display columns, row count, undo depth."""
    try:
        wb, ws, df = _read_active_sheet()
        app = wb.app

        if df.empty and not list(df.columns):
            _msgbox(app, 'No data found in the active sheet.', icon=48)
            return

        col_lines = '\n'.join(f'  {i+1:>2}. {c}' for i, c in enumerate(df.columns))
        msg = (f'Sheet : {ws.name}\n'
               f'Rows  : {len(df)}\n'
               f'Cols  : {len(df.columns)}\n\n'
               f'Columns:\n{col_lines}\n\n'
               f'Undo depth: {len(_undo_stack)} step(s)')
        _msgbox(app, msg, title='SheetMind — Sheet Info', icon=64)

    except Exception as e:
        try:
            xw.Book.caller().app.api.MsgBox(f'Error:\n{e}', 16, 'SheetMind')
        except Exception:
            pass


def _menu_row(parent, P, disp, body, icon, title, subtitle, cmd):
    """A sleek full-width action row (icon tile + title + subtitle)."""
    import tkinter as tk
    outer = tk.Frame(parent, bg=P['HAIR'])
    outer.pack(fill='x', pady=4)
    row = tk.Frame(outer, bg=P['CARD'], cursor='hand2')
    row.pack(fill='both', expand=True, padx=1, pady=1)

    tile = tk.Label(row, text=icon, bg=P['CARD2'], fg=P['ACCENT'],
                    font=(body, 15), width=2, pady=8)
    tile.pack(side='left', padx=(10, 12), pady=10)
    txt = tk.Frame(row, bg=P['CARD'])
    txt.pack(side='left', fill='x', expand=True, pady=8)
    tk.Label(txt, text=title, bg=P['CARD'], fg=P['TXT'], font=(disp, 12, 'bold'),
             anchor='w').pack(anchor='w')
    tk.Label(txt, text=subtitle, bg=P['CARD'], fg=P['SUB'], font=(body, 9),
             anchor='w').pack(anchor='w')
    chev = tk.Label(row, text='›', bg=P['CARD'], fg=P['DIM'], font=(disp, 16))
    chev.pack(side='right', padx=12)

    widgets = [row, txt, tile, chev] + list(txt.winfo_children())
    def on(_=None):
        for w in (row, txt, chev, *txt.winfo_children()):
            try: w['bg'] = P['CARD2']
            except Exception: pass
    def off(_=None):
        for w in (row, txt, chev, *txt.winfo_children()):
            try: w['bg'] = P['CARD']
            except Exception: pass
    for w in widgets:
        w.bind('<Enter>', on); w.bind('<Leave>', off); w.bind('<Button-1>', lambda e: cmd())
    return outer


def main():
    """Single-button entry point (xlwings fallback): a premium action menu
    matching the SheetMind web design — gradient header, sleek rows."""
    import tkinter as tk
    W = 380
    root, P, (disp, body, mono) = _make_window('SheetMind')
    chosen: list[str] = []

    _gradient_header(root, P, disp, mono, 'SheetMind', 'copilot for excel', W)

    wrap = tk.Frame(root, bg=P['BG'])
    wrap.pack(fill='both', expand=True, padx=18, pady=16)
    tk.Label(wrap, text='What would you like to do?', bg=P['BG'], fg=P['SUB'],
             font=(body, 10)).pack(anchor='w', pady=(0, 8))

    def pick(a):
        chosen.append(a); root.destroy()

    _menu_row(wrap, P, disp, body, '▶', 'Run a command',
              'Type a request — edit, sort, or ask a question', lambda: pick('text'))
    _menu_row(wrap, P, disp, body, '🎙', 'Voice command',
              'Speak; Whisper transcribes and runs it', lambda: pick('voice'))
    _menu_row(wrap, P, disp, body, '↩', 'Undo last change',
              'Revert the most recent edit', lambda: pick('undo'))
    _menu_row(wrap, P, disp, body, 'ℹ', 'Sheet info',
              'Columns, row count, undo depth', lambda: pick('info'))

    _center(root, W, 380)
    root.mainloop()

    if not chosen:
        return
    {'text': run_text_command, 'voice': run_voice_command,
     'undo': undo_last, 'info': show_info}[chosen[0]]()
