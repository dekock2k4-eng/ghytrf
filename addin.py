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


def _process_command(app, ws, df, cmd: str) -> None:
    """Shared pipeline for ribbon text/voice commands: interpret → preview →
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
        if not _confirm_box(app, _preview_text(plan, pv)):
            return
        done = engine.commit(plan, store)
        _msgbox(app, '✅ Applied.\n\n' + _format_results(done.steps), icon=64)
    else:
        done = engine.commit(plan, store)   # read-only — no save
        _msgbox(app, _format_results(done.steps), title='SheetMind — Result', icon=64)


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
            hint = (f'\n\nExamples for this sheet:\n'
                    f'  • Increase {cols[1]} by 10% where {cols[0]} is Example\n'
                    f'  • What is the total {cols[1]}?\n'
                    f'  • Show the top 5 rows by {cols[1]}\n'
                    f'  • Delete rows where {cols[1]} is less than 5')

        cmd = _inputbox(app, f'Enter your command:{hint}',
                        'SheetMind — Run Command')
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
        root = tk.Tk()
        root.title('SheetMind — Voice Command')
        root.resizable(False, False)
        root.attributes('-topmost', True)
        W, H = 400, 220
        root.update_idletasks()
        root.geometry(f'{W}x{H}+{(root.winfo_screenwidth()-W)//2}+{(root.winfo_screenheight()-H)//2}')

        recording_chunks: list = []
        is_recording = threading.Event()
        result = {'transcript': None}

        # UI elements
        tk.Label(root, text='SheetMind  Voice Command',
                 font=('Segoe UI', 13, 'bold')).pack(pady=(18, 4))
        status_var = tk.StringVar(value='Press Start to record from your microphone')
        status_lbl = tk.Label(root, textvariable=status_var,
                              font=('Segoe UI', 9), fg='#555555', wraplength=360)
        status_lbl.pack(pady=4)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=14)
        btn_text = tk.StringVar(value='  Start Recording  ')
        rec_btn = tk.Button(btn_frame, textvariable=btn_text,
                            font=('Segoe UI', 10, 'bold'), width=20, pady=7,
                            bg='#0078d4', fg='white', activebackground='#005a9e',
                            activeforeground='white', relief='flat', cursor='hand2')
        rec_btn.pack(side='left', padx=6)
        tk.Button(btn_frame, text='  Cancel  ',
                  font=('Segoe UI', 10), width=10, pady=7, relief='flat',
                  command=root.destroy, cursor='hand2').pack(side='left', padx=6)

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
                    status_lbl.config(fg='red'),
                    rec_btn.config(state='normal', bg='#0078d4', text='  Start Recording  '),
                    btn_text.set('  Start Recording  '),
                ))

        def toggle():
            if not is_recording.is_set():
                # START
                recording_chunks.clear()
                is_recording.set()
                threading.Thread(target=_record_worker, daemon=True).start()
                btn_text.set('  Stop & Transcribe  ')
                rec_btn.config(bg='#c42b1c', activebackground='#a01a0d')
                status_var.set('Recording...  Speak your command now')
                status_lbl.config(fg='#c42b1c')
            else:
                # STOP
                is_recording.clear()
                rec_btn.config(state='disabled')
                btn_text.set('  Transcribing...  ')
                status_var.set('Sending to Whisper — please wait...')
                status_lbl.config(fg='#0078d4')
                threading.Thread(target=_do_transcribe, daemon=True).start()

        rec_btn.config(command=toggle)
        root.mainloop()

        if result['transcript'] is None:
            return

        transcript = result['transcript']

        # ── Edit / confirm window ──────────────────────────────────────────
        root2 = tk.Tk()
        root2.title('SheetMind — Confirm Command')
        root2.resizable(False, False)
        root2.attributes('-topmost', True)
        W2, H2 = 460, 200
        root2.update_idletasks()
        root2.geometry(f'{W2}x{H2}+{(root2.winfo_screenwidth()-W2)//2}+{(root2.winfo_screenheight()-H2)//2}')

        final = {'cmd': None}

        tk.Label(root2, text='Whisper heard — edit if needed:',
                 font=('Segoe UI', 10, 'bold')).pack(pady=(18, 6))
        entry = tk.Entry(root2, font=('Segoe UI', 11), width=50, relief='solid', bd=1)
        entry.insert(0, transcript)
        entry.select_range(0, 'end')
        entry.pack(padx=20, ipady=5)
        entry.focus()

        def confirm():
            final['cmd'] = entry.get().strip()
            root2.destroy()

        f2 = tk.Frame(root2)
        f2.pack(pady=16)
        tk.Button(f2, text='  Run Command  ', command=confirm,
                  font=('Segoe UI', 10, 'bold'), pady=7, relief='flat',
                  bg='#0078d4', fg='white', activebackground='#005a9e',
                  activeforeground='white', cursor='hand2').pack(side='left', padx=6)
        tk.Button(f2, text='  Cancel  ', command=root2.destroy,
                  font=('Segoe UI', 10), pady=7, relief='flat',
                  cursor='hand2').pack(side='left', padx=6)
        root2.bind('<Return>', lambda _: confirm())
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


def main():
    """
    Single-button entry point used by the xlwings quickstart fallback.
    Shows a tkinter menu so the user can pick any SheetMind action.
    """
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title('SheetMind AI')
    root.resizable(False, False)
    root.attributes('-topmost', True)

    # Centre on screen
    root.update_idletasks()
    w, h = 320, 240
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f'{w}x{h}+{x}+{y}')

    chosen: list[str] = []

    tk.Label(root, text='SheetMind 🧠', font=('Segoe UI', 14, 'bold')).pack(pady=(18, 4))
    tk.Label(root, text='What would you like to do?',
             font=('Segoe UI', 9)).pack(pady=(0, 12))

    def pick(action: str) -> None:
        chosen.append(action)
        root.destroy()

    btn_cfg = dict(width=28, pady=6)
    tk.Button(root, text='▶  Run Text Command',  command=lambda: pick('text'),  **btn_cfg).pack(pady=3)
    tk.Button(root, text='🎙  Voice Command',     command=lambda: pick('voice'), **btn_cfg).pack(pady=3)
    tk.Button(root, text='↩  Undo Last Change',  command=lambda: pick('undo'),  **btn_cfg).pack(pady=3)
    tk.Button(root, text='ℹ  Sheet Info',         command=lambda: pick('info'),  **btn_cfg).pack(pady=3)

    root.mainloop()

    if not chosen:
        return

    action = chosen[0]
    if action == 'text':
        run_text_command()
    elif action == 'voice':
        run_voice_command()
    elif action == 'undo':
        undo_last()
    elif action == 'info':
        show_info()
