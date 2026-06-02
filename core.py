"""
core.py — Final Clean & Stable Version
"""

from __future__ import annotations
import io
import os
import tempfile
import json
import re
import uuid
import time
from datetime import datetime
from typing import Optional

import pandas as pd
import openpyxl
from openpyxl import Workbook
from dotenv import load_dotenv

load_dotenv()

def _resolve_save_dir() -> str:
    env = os.environ.get("SHEETMIND_SAVE_DIR", "").strip()
    if env:
        os.makedirs(env, exist_ok=True)
        return env
    script = os.path.dirname(os.path.abspath(__file__))
    try:
        probe = os.path.join(script, ".write_probe")
        open(probe, "w").close()
        os.unlink(probe)
        return script
    except OSError:
        pass
    fb = os.path.join(tempfile.gettempdir(), "sheetmind_files")
    os.makedirs(fb, exist_ok=True)
    return fb

_SAVE_DIR = _resolve_save_dir()
_UNDO_LIMIT = 25

_GEMINI_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.0-flash"]


class SheetStore:
    """Main data layer for SheetMind."""

    def __init__(self):
        self.file_path: Optional[str] = None
        self.file_name: Optional[str] = None
        self.active_sheet: Optional[str] = None
        self.sheet_names: list[str] = []
        self._df: Optional[pd.DataFrame] = None
        self._history: list = []
        self._file_bytes: Optional[bytes] = None

    @property
    def df(self) -> pd.DataFrame:
        return self._df if self._df is not None else pd.DataFrame()

    @property
    def file_bytes(self) -> Optional[bytes]:
        """Return cached bytes or read from disk if missing."""
        if self._file_bytes is None and self.file_path and os.path.exists(self.file_path):
            with open(self.file_path, "rb") as f:
                self._file_bytes = f.read()
        return self._file_bytes

    @property
    def is_loaded(self) -> bool:
        return self.file_path is not None and self._df is not None

    @property
    def undo_depth(self) -> int:
        return len(self._history)

    # ====================== LOAD ======================
    def load_from_upload(self, raw: bytes, file_name: str) -> "SheetStore":
        self._history = []
        self.sheet_names = []
        self.active_sheet = None
        self._df = None
        self._file_bytes = None
        safe_name = os.path.basename(file_name).strip()
        path = _get_save_path(safe_name)

        with open(path, "wb") as f:
            f.write(raw)

        return self._load(path, os.path.basename(path))

    def _load(self, path: str, name: str) -> "SheetStore":
        self.file_path = path
        self.file_name = name
        self._history = []

        engine = _get_excel_engine(path)
        try:
            with pd.ExcelFile(path, engine=engine) as xl:
                self.sheet_names = xl.sheet_names
                self.active_sheet = xl.sheet_names[0]
        except ImportError as e:
            if engine == "xlrd":
                raise RuntimeError(
                    "Loading .xls files requires xlrd. Convert the file to .xlsx "
                    "or install a compatible xlrd version."
                ) from e
            raise
        except Exception as e:
            raise RuntimeError(f"Failed to open Excel file: {e}") from e

        self._read_sheet(self.active_sheet, engine=engine)
        self._cache_bytes()
        return self

    def _read_sheet(self, sheet_name: str, engine: Optional[str] = None):
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine=engine)
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.columns:
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().sum() > 0 and converted.notna().mean() > 0.5:
                df[col] = converted
        self._df = df

    def _cache_bytes(self):
        if self.file_path and os.path.exists(self.file_path):
            with open(self.file_path, "rb") as f:
                self._file_bytes = f.read()

    # ====================== SAVE ======================
    def save(self) -> None:
        """
        Write DataFrame to disk atomically.
        Strategy:
          1. Write to a .tmp file first (never corrupts the original if killed mid-write)
          2. Use os.replace() for atomic rename (POSIX) or shutil.move fallback (Windows lock)
          3. On Windows, if the original file is open in Excel, fall back to writing
             directly so the user at least gets *a* save attempt
          4. Always raises RuntimeError on failure — callers must NOT silently ignore it
        """
        if not self.file_path or self._df is None:
            raise RuntimeError("SheetStore: nothing to save (file_path or _df is None)")

        # Convert .xls tracked paths to .xlsx before writing,
        # otherwise openpyxl would write XLSX bytes into an .xls file.
        original_path = self.file_path
        if original_path.lower().endswith(".xls"):
            self.file_path = os.path.splitext(original_path)[0] + ".xlsx"
            self.file_name = os.path.basename(self.file_path)

        import shutil as _shutil

        # Write to in-memory buffer first — avoids any file-locking issues during write
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            sheet_name = self.active_sheet or "Sheet1"
            self._df.to_excel(writer, index=False, sheet_name=sheet_name)
        xlsx_bytes = buf.getvalue()

        # Try atomic write via temp file + rename
        tmp_path = self.file_path + ".sheetmind_tmp"
        try:
            with open(tmp_path, "wb") as f:
                f.write(xlsx_bytes)
            try:
                os.replace(tmp_path, self.file_path)
            except OSError:
                # Windows: file may be open in Excel — os.replace fails with PermissionError.
                # Fall back to shutil.move which tries harder.
                try:
                    _shutil.move(tmp_path, self.file_path)
                except Exception:
                    # Last resort: write directly (no atomicity, but still saves)
                    with open(self.file_path, "wb") as f:
                        f.write(xlsx_bytes)
                    if os.path.exists(tmp_path):
                        try: os.unlink(tmp_path)
                        except: pass
        except Exception as e:
            if os.path.exists(tmp_path):
                try: os.unlink(tmp_path)
                except: pass
            raise RuntimeError(f"Save failed for {self.file_path!r}: {e}") from e

        # Update in-memory cache so download button is always in sync
        self._file_bytes = xlsx_bytes

        if original_path != self.file_path and os.path.exists(original_path):
            try:
                os.unlink(original_path)
            except:
                pass

    # ====================== UNDO ======================
    def _push_history(self):
        if self._df is not None:
            self._history.append((self.active_sheet, self._df.copy(deep=True)))
            if len(self._history) > _UNDO_LIMIT:
                self._history.pop(0)

    def undo(self) -> bool:
        if not self._history:
            return False
        sheet, snap = self._history.pop()
        self.active_sheet = sheet
        self._df = snap
        self.save()
        return True

    # ====================== CRUD ======================
    def add_row(self, data: dict):
        self._push_history()
        row = {c: data.get(c, "") for c in self._df.columns}
        self._df = pd.concat([self._df, pd.DataFrame([row])], ignore_index=True)
        self.save()

    def update_row(self, idx: int, data: dict):
        self._push_history()
        df = self._df.copy()
        for col, val in data.items():
            if col in df.columns:
                df.at[idx, col] = val
        self._df = df
        self.save()

    def delete_rows(self, indices: list):
        self._push_history()
        self._df = self._df.drop(index=indices).reset_index(drop=True)
        self.save()

    def delete_rows_matching(self, col: str, value: str) -> int:
        if col not in self._df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(self._df.columns)}")
        self._push_history()
        mask  = self._df[col].astype(str).str.strip().str.lower() == str(value).strip().lower()
        count = int(mask.sum())
        self._df = self._df[~mask].reset_index(drop=True)
        self.save()
        return count

    def update_cells_matching(self, match_col: str, match_val: str, target_col: str, new_val) -> int:
        if match_col not in self._df.columns:
            raise ValueError(f"Column '{match_col}' not found. Available: {list(self._df.columns)}")
        if target_col not in self._df.columns:
            raise ValueError(f"Column '{target_col}' not found. Available: {list(self._df.columns)}")
        self._push_history()
        df   = self._df.copy()
        mask = df[match_col].astype(str).str.strip().str.lower() == str(match_val).strip().lower()

        if isinstance(new_val, str) and new_val.startswith(('+', '-')):
            try:
                delta = float(new_val)
                current = pd.to_numeric(df.loc[mask, target_col], errors="coerce").fillna(0)
                df.loc[mask, target_col] = current + delta
            except:
                df.loc[mask, target_col] = new_val
        else:
            df.loc[mask, target_col] = new_val

        self._df = df
        self.save()
        return int(mask.sum())

    def clear_rows(self):
        self._push_history()
        self._df = pd.DataFrame(columns=self._df.columns)
        self.save()

    def add_column(self, name: str):
        if name in self._df.columns:
            raise ValueError(f"Column '{name}' already exists")
        self._push_history()
        self._df[name] = ""
        self.save()

    def delete_column(self, name: str):
        if name not in self._df.columns:
            raise ValueError(f"Column '{name}' not found")
        self._push_history()
        self._df = self._df.drop(columns=[name])
        self.save()

    def rename_column(self, old: str, new: str):
        if old not in self._df.columns:
            raise ValueError(f"Column '{old}' not found")
        if new in self._df.columns:
            raise ValueError(f"Column '{new}' already exists")
        self._push_history()
        self._df = self._df.rename(columns={old: new})
        self.save()


    @classmethod
    def create_blank(cls, file_name: str, columns: list) -> "SheetStore":
        safe = os.path.basename(file_name.strip())
        path = _get_save_path(safe)
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(columns)
        wb.save(path)
        store = cls()
        store.file_path    = path
        store.file_name    = safe
        store.sheet_names  = ["Sheet1"]
        store.active_sheet = "Sheet1"
        store._df          = pd.DataFrame(columns=columns)
        store._history     = []
        store._cache_bytes()
        return store

    def load_from_path(self, path: str) -> "SheetStore":
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        return self._load(path, os.path.basename(path))

    def switch_sheet(self, name: str) -> None:
        if name in self.sheet_names and name != self.active_sheet:
            self._push_history()
            self.active_sheet = name
            self._read_sheet(name)


# ====================== TRANSCRIPTION ======================
def transcribe(audio_bytes: bytes) -> str:
    """
    Transcribe spoken audio to text via Whisper large-v3-turbo on Hugging Face.

    Multilingual by default — Whisper auto-detects the spoken language and the
    downstream LLM command engine (Kimi K2) is itself multilingual, so a Hindi or
    Spanish command works end-to-end. Set SHEETMIND_FORCE_LANGUAGE (e.g. "en") to
    pin a language; leave unset for auto-detect.
    """
    token = os.environ.get("HF_TOKEN")
    if not token:
        raise ValueError(
            "HF_TOKEN missing from .env — voice needs a free Hugging Face token "
            "(https://huggingface.co/settings/tokens). You can still type commands."
        )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
    tmp.write(audio_bytes)
    tmp.close()

    lang = os.environ.get("SHEETMIND_FORCE_LANGUAGE", "").strip() or None

    try:
        from huggingface_hub import InferenceClient
        client = InferenceClient(provider="hf-inference", api_key=token)
        model  = "openai/whisper-large-v3-turbo"
        result = None
        errors = []
        params = {"task": "transcribe"}
        if lang:
            params["language"] = lang
        # Attempt 1: parameters= (current HF API)
        try:
            result = client.automatic_speech_recognition(
                audio=tmp.name, model=model, parameters=params,
            )
        except Exception as e:
            errors.append(str(e))
        # Attempt 2: generate_kwargs= (older HF client)
        if result is None:
            try:
                result = client.automatic_speech_recognition(
                    audio=tmp.name, model=model, generate_kwargs=params,
                )
            except Exception as e:
                errors.append(str(e))
        # Attempt 3: plain call (let Whisper auto-detect everything)
        if result is None:
            try:
                result = client.automatic_speech_recognition(audio=tmp.name, model=model)
            except Exception as e:
                errors.append(str(e))
        if result is None:
            raise RuntimeError("Whisper failed: " + " | ".join(errors))
        return (result.text if hasattr(result, "text") else str(result)).strip()
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


# ====================== COMMAND PARSER ======================
_PROMPT = """\
You are a strict spreadsheet command parser. Output ONLY a JSON object — no prose, no markdown.

Sheet columns (use EXACT names): {cols}
Sample rows: {sample}

User command: "{cmd}"

Return one of:
{{"action":"add_row",    "data":{{"ColName":"value"}}}}
{{"action":"set_cell",   "match_col":"ColName","match_val":"val","target_col":"ColName","value":"newval"}}
{{"action":"update_row", "match_col":"ColName","match_val":"val","updates":{{"ColName":"val"}}}}
{{"action":"delete_row", "match_col":"ColName","match_val":"val"}}

Rules:
- Copy column names EXACTLY from the list above.
- Arithmetic (add 5, increase by 10) -> set_cell with value "+5" or "+10".
- Subtraction -> value "-5".
- First character of response MUST be {{.
"""

_ONES = {
    'zero':0,'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,
    'eight':8,'nine':9,'ten':10,'eleven':11,'twelve':12,'thirteen':13,
    'fourteen':14,'fifteen':15,'sixteen':16,'seventeen':17,'eighteen':18,
    'nineteen':19,'twenty':20,'thirty':30,'forty':40,'fifty':50,'sixty':60,
    'seventy':70,'eighty':80,'ninety':90,
}
_MULTIPLIERS = {'hundred':100,'thousand':1_000,'million':1_000_000,'billion':1_000_000_000}
_NUMBER_WORDS = set(_ONES) | set(_MULTIPLIERS) | {'and', 'a'}

def _word_to_number(text: str):
    """
    Convert English number words to int or float.
    Returns original string unchanged if the text is not a number phrase.
    Examples: 'fifty thousand' -> 50000, 'one hundred twenty' -> 120
    """
    if not isinstance(text, str):
        return text
    s = text.strip().lower().replace('-', ' ').replace(',', '')
    words = s.split()
    if not words:
        return text
    # Only convert if every word is a recognised number word
    if not all(w in _NUMBER_WORDS for w in words):
        return text
    current = 0
    result = 0
    for word in words:
        if word in ('and', 'a'):
            continue
        if word in _ONES:
            current += _ONES[word]
        elif word == 'hundred':
            current = (current if current else 1) * 100
        elif word in _MULTIPLIERS:
            result += (current if current else 1) * _MULTIPLIERS[word]
            current = 0
    result += current
    return result


def _coerce_value(val):
    """
    Try to convert a value to a number.
    Priority: already numeric → word-number phrase → numeric string → original.
    """
    if isinstance(val, (int, float)):
        return val
    if not isinstance(val, str):
        return val
    # Skip arithmetic deltas like "+50" or "-10" — handled separately
    if re.match(r'^[+-][\d.]+$', val.strip()):
        return val
    # Try word-to-number conversion first
    converted = _word_to_number(val)
    if converted != val:
        return converted
    # Try plain numeric string
    try:
        f = float(val)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        pass
    return val


def _fuzzy_col(name: str, cols: list) -> str:
    nl = name.strip().lower()
    for c in cols:
        if c.lower() == nl: return c
    for c in cols:
        if nl in c.lower() or c.lower() in nl: return c
    return name

def _local_parse(text: str, cols: list):
    tl = text.strip().lower()
    # Arithmetic: "add 50 to Qty where Item is Pen"
    m = re.search(r'(add|increase|plus|subtract|decrease|minus)\s+([\d.]+)\s+(?:to|from)\s+(\w[\w\s]*?)\s+where\s+(\w[\w\s]*?)\s+(?:is|=)\s+(.+)', tl)
    if m:
        op = m.group(1); amt = float(m.group(2))
        tc = _fuzzy_col(m.group(3).strip(), cols)
        mc = _fuzzy_col(m.group(4).strip(), cols)
        mv = m.group(5).strip().strip('"\' ')
        delta = amt if op in ("add","increase","plus") else -amt
        if tc in cols and mc in cols:
            # Format delta: strip trailing .0 for whole numbers (50.0 → 50)
            def _fmt(n):
                return str(int(n)) if n == int(n) else str(n)
            val_str = f"+{_fmt(delta)}" if delta >= 0 else _fmt(delta)
            return {"action":"set_cell","match_col":mc,"match_val":mv,
                    "target_col":tc,"value":val_str}, ""
    # Set: "set Price to 25 where Item is Pen"
    m = re.search(r'(?:set|update|change)\s+(\w[\w\s]*?)\s+to\s+(.+?)\s+where\s+(\w[\w\s]*?)\s+(?:is|=)\s+(.+)', tl)
    if m:
        tc = _fuzzy_col(m.group(1).strip(), cols); val = m.group(2).strip().strip('"\' ')
        mc = _fuzzy_col(m.group(3).strip(), cols); mv  = m.group(4).strip().strip('"\' ')
        if tc in cols and mc in cols:
            return {"action":"set_cell","match_col":mc,"match_val":mv,"target_col":tc,"value":val}, ""
    # Delete: "delete row where Item is Pen"
    m = re.search(r'(?:delete|remove)\s+(?:row\s+)?(?:where\s+)?(\w[\w\s]*?)\s+(?:is|=)\s+(.+)', tl)
    if m:
        mc = _fuzzy_col(m.group(1).strip(), cols); mv = m.group(2).strip().strip('"\' ')
        if mc in cols:
            return {"action":"delete_row","match_col":mc,"match_val":mv}, ""
    return None, ""

def parse_command(text: str, store: SheetStore) -> tuple[dict, str]:
    text = text.strip()
    if not text: return {}, "Empty command."
    if len(text) > 500: return {}, "Command too long (max 500 chars)."

    cols   = list(store.df.columns)
    sample = store.df.head(3).to_dict(orient="records")

    local, _ = _local_parse(text, cols)
    if local: return local, ""

    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not key:
        return {}, "GEMINI_API_KEY missing from .env"

    try:
        from google import genai
    except ImportError:
        return {}, "google-genai SDK not installed."

    client   = genai.Client(api_key=key)
    prompt   = _PROMPT.format(cols=cols, sample=json.dumps(sample[:3], default=str), cmd=text)
    last_err = "no models tried"

    for model in _GEMINI_MODELS:
        for attempt in range(2):
            try:
                resp = client.models.generate_content(model=model, contents=prompt)
                raw  = (resp.text or "").strip()
                fb = raw.find("{"); lb = raw.rfind("}")
                if fb != -1 and lb > fb:
                    raw = raw[fb:lb+1]
                obj = json.loads(raw)
                if isinstance(obj, list): obj = obj[0]
                if "action" not in obj: raise ValueError("Missing 'action' field")
                for field in ("match_col", "target_col"):
                    v = obj.get(field)
                    if v and v not in cols:
                        obj[field] = _fuzzy_col(v, cols)
                return obj, ""
            except Exception as e:
                last_err = str(e)
                if any(s in last_err for s in ("NOT_FOUND","404","not found","not supported")):
                    break
                if any(s in last_err for s in ("503","UNAVAILABLE")):
                    time.sleep(2**attempt); continue
                if any(s in last_err for s in ("429","RESOURCE_EXHAUSTED","quota")):
                    break
                break

    return {}, f"Gemini failed: {last_err}"


def apply_command(action: dict, store: SheetStore) -> str:
    """
    Apply a parsed action to SheetStore.
    RuntimeError (save failures) are re-raised so the UI can show the correct message.
    ValueError / KeyError (bad column names etc.) are returned as ❌ strings.
    """
    act = action.get("action", "")
    try:
        if act == "add_row":
            data = action.get("data", {})
            if not data:
                return "⚠️ add_row: no data provided."
            data = {k: _coerce_value(v) for k, v in data.items()}
            store.add_row(data)
            return f"✅ Row added: {data}"

        elif act == "set_cell":
            mc  = action.get("match_col")
            mv  = action.get("match_val")
            tc  = action.get("target_col")
            val = action.get("value")
            if not all([mc, mv is not None, tc, val is not None]):
                return "⚠️ set_cell: missing match_col, match_val, target_col or value."
            val = _coerce_value(val)
            count = store.update_cells_matching(mc, str(mv), tc, val)
            if count == 0:
                return f"⚠️ No rows found where {mc} = {mv}"
            return f"✅ Updated {count} cell(s): {tc} = {val}"

        elif act == "update_row":
            mc  = action.get("match_col")
            mv  = action.get("match_val")
            upd = action.get("updates", {})
            if not mc or mv is None or not upd:
                return "⚠️ update_row: missing match_col, match_val or updates."
            df   = store.df
            mask = df[mc].astype(str).str.strip().str.lower() == str(mv).strip().lower()
            if not mask.any():
                return f"⚠️ No rows found where {mc} = {mv}"
            store._push_history()
            df2 = df.copy()
            for col, v in upd.items():
                if col not in df2.columns:
                    continue
                v = _coerce_value(v)
                vs = str(v)
                if re.match(r'^[+-][\d.]+$', vs.strip()):
                    delta = float(vs)
                    df2.loc[mask, col] = (
                        pd.to_numeric(df2.loc[mask, col], errors="coerce").fillna(0) + delta
                    )
                else:
                    df2.loc[mask, col] = v
            store._df = df2
            store.save()
            return f"✅ Updated {int(mask.sum())} row(s)"

        elif act == "delete_row":
            mc = action.get("match_col")
            mv = action.get("match_val")
            if not mc or mv is None:
                return "⚠️ delete_row: missing match_col or match_val."
            count = store.delete_rows_matching(mc, str(mv))
            if count == 0:
                return f"⚠️ No rows found where {mc} = {mv}"
            return f"✅ Deleted {count} row(s) where {mc} = {mv}"

        else:
            return f"⚠️ Unknown action: '{act}'. Expected: add_row, set_cell, update_row, delete_row."

    except RuntimeError:
        raise   # Save failures must reach the UI — don't swallow them
    except (ValueError, KeyError) as e:
        return f"❌ {e}"
    except Exception as e:
        return f"❌ Unexpected error: {e}"

def _get_excel_engine(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        return "xlrd"
    return "openpyxl"

def _get_save_path(file_name: str) -> str:
    name, ext = os.path.splitext(os.path.basename(file_name).strip())
    ext = ext.lower()
    if ext not in (".xlsx", ".xls"):
        ext = ".xlsx"
    return os.path.join(_SAVE_DIR, f"{name}{ext}")