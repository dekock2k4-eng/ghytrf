"""
engine.py — SheetMind's natural-language command engine.

This is the brain that turns a sentence ("bump every electronics item's price by
10% and show me the new total") into a validated, previewable PLAN of typed
operations, executed deterministically against a pandas DataFrame.

Design goals (what makes it "world-class"):
  • Rich vocabulary  — filters with real operators (>, <, contains, between, …),
                       sort, computed/formula columns, multi-column updates,
                       and read-only analytics (aggregate / top-N) for Q&A.
  • Grounded         — every plan is built against the live sheet schema + stats.
  • Tool-calling     — the LLM returns a structured plan via a forced function
                       call (engine never parses JSON out of prose).
  • Preview → confirm — mutations are dry-run on a copy first; the UI shows a diff
                       and only commits on explicit confirmation.
  • Deterministic math — all numbers in answers are computed by pandas here, never
                       hallucinated by the model.

The engine is store-agnostic: it works on any object exposing ``.df``,
``_push_history()`` and ``save()`` (both SheetStore and the Excel XlwingsStore).
"""
from __future__ import annotations

import re
import json
from dataclasses import dataclass, field
from typing import Any, Optional

import numpy as np
import pandas as pd

from llm import chat, is_configured, LLMNotConfigured

# Re-use the existing word→number + fuzzy-column helpers from core.
from core import _word_to_number, _coerce_value, _fuzzy_col


# ════════════════════════════════════════════════════════════════════════
# Data structures
# ════════════════════════════════════════════════════════════════════════
MUTATING_OPS = {
    "add_row", "update", "delete", "clear_rows",
    "add_column", "delete_column", "rename_column", "sort",
}
READ_OPS = {"aggregate", "top_n", "describe"}


@dataclass
class OpResult:
    op: str
    kind: str               # "mutate" | "read"
    summary: str            # human-readable one-liner
    affected: int = 0       # rows/cells/cols touched (mutations)
    table: Optional[pd.DataFrame] = None   # read results to render
    value: Any = None       # scalar read result
    error: Optional[str] = None


@dataclass
class Plan:
    operations: list[dict]
    summary: str = ""
    source: str = "llm"     # "llm" | "local"
    raw: dict = field(default_factory=dict)

    @property
    def has_mutation(self) -> bool:
        return any(o.get("op") in MUTATING_OPS for o in self.operations)

    @property
    def needs_confirm(self) -> bool:
        return self.has_mutation


@dataclass
class PreviewResult:
    plan: Plan
    steps: list[OpResult]
    before: Optional[pd.DataFrame] = None    # snapshot before mutations
    after: Optional[pd.DataFrame] = None     # snapshot after mutations (dry-run)
    error: Optional[str] = None

    @property
    def ok(self) -> bool:
        return self.error is None and all(s.error is None for s in self.steps)


# ════════════════════════════════════════════════════════════════════════
# Condition / filter evaluation
# ════════════════════════════════════════════════════════════════════════
def _to_num(v):
    """Best-effort numeric coercion of a scalar (handles word-numbers)."""
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        w = _word_to_number(v)
        if isinstance(w, (int, float)):
            return w
        try:
            return float(v)
        except (ValueError, TypeError):
            return None
    return None


def _clean_num(val):
    """Normalise a numpy/pandas scalar to a plain int/float for display."""
    if isinstance(val, (np.floating, float)):
        f = float(val)
        return int(f) if f == int(f) else round(f, 4)
    if isinstance(val, (np.integer,)):
        return int(val)
    return val


def _agg_series(series: pd.Series, func: str):
    return {"sum": series.sum(), "mean": series.mean(), "avg": series.mean(),
            "min": series.min(), "max": series.max(), "median": series.median(),
            "std": series.std(), "count": series.count()}.get(func, series.sum())


def _infer_numeric_col(df: pd.DataFrame) -> str:
    """Best-effort single numeric column (used when the model omits 'column')."""
    numeric = [c for c in df.columns if pd.to_numeric(df[c], errors="coerce").notna().any()]
    return numeric[0] if numeric else ""


def _cond_mask(df: pd.DataFrame, cond: dict) -> pd.Series:
    cols = list(df.columns)
    col = _fuzzy_col(str(cond.get("column", "")), cols)
    if col not in df.columns:
        raise ValueError(f"Column '{cond.get('column')}' not found. Available: {cols}")
    op = str(cond.get("op", "eq")).lower()
    val = cond.get("value")
    series = df[col]

    if op in ("gt", "gte", "lt", "lte", "between"):
        s = pd.to_numeric(series, errors="coerce")
        if op == "between":
            if not isinstance(val, (list, tuple)) or len(val) != 2:
                raise ValueError("'between' needs a [low, high] value")
            lo, hi = _to_num(val[0]), _to_num(val[1])
            return (s >= lo) & (s <= hi)
        v = _to_num(val)
        if v is None:
            raise ValueError(f"'{op}' needs a numeric value, got {val!r}")
        return {"gt": s > v, "gte": s >= v, "lt": s < v, "lte": s <= v}[op]

    if op in ("contains", "startswith", "endswith"):
        s = series.astype(str).str.strip().str.lower()
        needle = str(val).strip().lower()
        return {
            "contains": s.str.contains(re.escape(needle), na=False),
            "startswith": s.str.startswith(needle, na=False),
            "endswith": s.str.endswith(needle, na=False),
        }[op]

    if op in ("isblank", "isempty"):
        return series.isna() | (series.astype(str).str.strip() == "")
    if op in ("notblank", "notempty"):
        return ~(series.isna() | (series.astype(str).str.strip() == ""))

    if op == "in":
        vals = val if isinstance(val, (list, tuple)) else [val]
        wanted = {str(x).strip().lower() for x in vals}
        return series.astype(str).str.strip().str.lower().isin(wanted)

    # eq / ne — prefer numeric comparison when both sides look numeric.
    nv = _to_num(val)
    if nv is not None:
        sn = pd.to_numeric(series, errors="coerce")
        if sn.notna().any():
            eq = sn == nv
            return eq if op == "eq" else ~eq
    eq = series.astype(str).str.strip().str.lower() == str(val).strip().lower()
    return eq if op == "eq" else ~eq


def _build_mask(df: pd.DataFrame, filt: Optional[dict]) -> pd.Series:
    if not filt or not filt.get("conditions"):
        return pd.Series([True] * len(df), index=df.index)
    match = str(filt.get("match", "all")).lower()
    masks = [_cond_mask(df, c) for c in filt["conditions"]]
    combined = masks[0]
    for m in masks[1:]:
        combined = (combined & m) if match == "all" else (combined | m)
    return combined


def _describe_filter(filt: Optional[dict]) -> str:
    if not filt or not filt.get("conditions"):
        return "all rows"
    sym = {"eq": "=", "ne": "≠", "gt": ">", "gte": "≥", "lt": "<", "lte": "≤",
           "contains": "contains", "startswith": "starts with", "endswith": "ends with",
           "between": "between", "in": "in", "isblank": "is blank", "notblank": "is not blank"}
    parts = []
    for c in filt["conditions"]:
        o = c.get("op", "eq")
        if o in ("isblank", "notblank"):
            parts.append(f"{c.get('column')} {sym.get(o, o)}")
        else:
            parts.append(f"{c.get('column')} {sym.get(o, o)} {c.get('value')}")
    joiner = " AND " if str(filt.get("match", "all")).lower() == "all" else " OR "
    return joiner.join(parts)


# ════════════════════════════════════════════════════════════════════════
# Safe formula evaluation  (columns referenced as [Col Name])
# ════════════════════════════════════════════════════════════════════════
_FORMULA_ALLOWED = re.compile(r"^[\s\d\.\+\-\*\/\%\(\)\,a-zA-Z_]*$")


def _eval_formula(formula: str, df: pd.DataFrame) -> pd.Series:
    """
    Evaluate a vectorised formula over the DataFrame. Columns are referenced
    with square brackets, e.g.  "[Qty] * [Price] * 1.1".  Supported: + - * / %
    ** ( ) and the functions round/abs/min/max. No attribute access, no calls
    to anything else — safe to run on user/LLM input.
    """
    refs = list(dict.fromkeys(re.findall(r"\[([^\]]+)\]", formula)))
    namespace: dict[str, Any] = {}
    expr = formula
    for i, ref in enumerate(refs):
        col = _fuzzy_col(ref, list(df.columns))
        if col not in df.columns:
            raise ValueError(f"Formula references unknown column '{ref}'")
        var = f"_c{i}_"
        col_data = pd.to_numeric(df[col], errors="coerce")
        namespace[var] = col_data if col_data.notna().any() else df[col]
        expr = expr.replace(f"[{ref}]", var)

    if not _FORMULA_ALLOWED.match(expr):
        raise ValueError(f"Formula contains unsupported characters: {formula!r}")

    # Vectorised math functions, exposed in BOTH Excel style (ROUND) and python
    # style (round) so either phrasing the model emits just works.
    _fns = {
        "round": np.round, "abs": np.abs, "min": np.minimum, "max": np.maximum,
        "int": np.trunc, "floor": np.floor, "ceil": np.ceil, "ceiling": np.ceil,
        "sqrt": np.sqrt, "sum": np.add, "trunc": np.trunc, "mod": np.mod,
    }
    safe_funcs = {"np": np}
    for k, v in _fns.items():
        safe_funcs[k] = v
        safe_funcs[k.upper()] = v
    try:
        result = eval(expr, {"__builtins__": {}}, {**namespace, **safe_funcs})  # noqa: S307
    except Exception as e:
        raise ValueError(f"Could not evaluate formula {formula!r}: {e}") from e
    if np.isscalar(result):
        result = pd.Series([result] * len(df), index=df.index)
    return result


def _sanitize_formulas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Safety net: ensure no internal "=[Col]" formula DSL ever reaches the saved
    file. Any such leftover string cell is evaluated to its real value (Excel
    would otherwise treat the leading '=' as a broken formula → #REF!). Plain
    Excel formulas like "=B2*C2" (no [brackets]) are left untouched on purpose.
    """
    out = None
    for col in df.columns:
        # Numeric/bool columns can't hold a formula string — skip for speed.
        if pd.api.types.is_numeric_dtype(df[col]) or pd.api.types.is_bool_dtype(df[col]):
            continue
        # Find offending cells in this column.
        hits = [idx for idx in df.index
                if isinstance(df.at[idx, col], str)
                and df.at[idx, col].strip().startswith("=")
                and "[" in df.at[idx, col] and "]" in df.at[idx, col]]
        if not hits:
            continue
        if out is None:
            out = df.copy()
        # Cast to object so we can write computed numbers into a string column.
        col_obj = out[col].astype(object)
        for idx in hits:
            v = out.at[idx, col]
            try:
                col_obj.at[idx] = _clean_num(_eval_formula(v.strip()[1:], out).loc[idx])
            except Exception:  # noqa: BLE001 — never let a bad formula block a save
                col_obj.at[idx] = ""
        # Restore numeric typing where the column is now fully numeric.
        try:
            out[col] = pd.to_numeric(col_obj)
        except (ValueError, TypeError):
            out[col] = col_obj
    return out if out is not None else df


def _resolve_set_value(raw, df: pd.DataFrame, mask: pd.Series, col: str):
    """
    Turn a `set` value into the concrete data to write into df.loc[mask, col].
    Supports: formula ("=[Qty]*[Price]"), arithmetic delta ("+5"/"-10"),
    word-numbers, or a literal.
    """
    if isinstance(raw, str):
        s = raw.strip()
        if s.startswith("="):
            return _eval_formula(s[1:], df)[mask]
        if re.match(r"^[+-][\d.]+$", s):
            delta = float(s)
            current = pd.to_numeric(df.loc[mask, col], errors="coerce").fillna(0)
            return current + delta
    return _coerce_value(raw)


def _safe_assign(df: pd.DataFrame, mask, col: str, value) -> None:
    """
    Write `value` into df.loc[mask, col], widening the column's dtype first when
    the incoming values don't fit it.

    Pandas 2.x raises `Invalid value '[...]' for dtype 'int64'` when you set
    fractional floats into an integer column — e.g. a 10% price bump turns
    Price [2,8,4,6] into [2.2,8.8,4.4,6.6]. We upcast the column to float (or
    object as a last resort for non-numeric values) and retry so the edit
    succeeds instead of erroring out.
    """
    try:
        df.loc[mask, col] = value
        return
    except (TypeError, ValueError):
        pass
    widen = "float64" if pd.api.types.is_integer_dtype(df[col]) else object
    df[col] = df[col].astype(widen)
    try:
        df.loc[mask, col] = value
        return
    except (TypeError, ValueError):
        df[col] = df[col].astype(object)
        df.loc[mask, col] = value


# ════════════════════════════════════════════════════════════════════════
# Operation executor  (pure: takes df, returns new df + OpResult)
# ════════════════════════════════════════════════════════════════════════
def _exec_op(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, OpResult]:
    name = op.get("op", "")
    cols = list(df.columns)

    # ── add_row ──────────────────────────────────────────────────────────
    if name == "add_row":
        values = op.get("values") or op.get("data") or {}
        if not values:
            return df, OpResult(name, "mutate", "add_row: no values", error="no values provided")
        row = {c: "" for c in cols}
        formulas = {}          # columns whose value is a "=[..]" formula
        for k, v in values.items():
            col = _fuzzy_col(k, cols)
            if isinstance(v, str) and v.strip().startswith("="):
                formulas[col] = v.strip()[1:]   # defer until the row exists
            else:
                row[col] = _coerce_value(v)
        df2 = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        # Evaluate any formula values against the freshly-added row so the cell
        # holds a real number — NOT the literal "=[Qty]*[Price]" (Excel would
        # treat that as a broken formula and show #REF!).
        for col, expr in formulas.items():
            try:
                df2.loc[df2.index[-1], col] = _eval_formula(expr, df2).iloc[-1]
            except ValueError as e:
                return df, OpResult(name, "mutate", f"add_row: bad formula for {col}",
                                    error=str(e))
        return df2, OpResult(name, "mutate", f"Add 1 row: {values}", affected=1)

    # ── update / set ─────────────────────────────────────────────────────
    if name in ("update", "set", "set_cell"):
        filt = op.get("filter")
        setmap = op.get("set") or op.get("updates") or {}
        # Back-compat / tolerant shapes: {target_col,value} or {column,formula}
        if not setmap:
            tcol = op.get("target_col") or op.get("column")
            if tcol:
                v = op.get("formula")
                if v is not None and not str(v).startswith("="):
                    v = "=" + str(v)
                if v is None:
                    v = op.get("value")
                if v is not None:
                    setmap = {tcol: v}
        if not setmap:
            return df, OpResult(name, "mutate", "update: nothing to set", error="no 'set' provided")
        # Back-compat single match shape
        if filt is None and op.get("match_col"):
            filt = {"match": "all", "conditions":
                    [{"column": op["match_col"], "op": "eq", "value": op.get("match_val")}]}
        mask = _build_mask(df, filt)
        n = int(mask.sum())
        df2 = df.copy()
        for col_raw, val in setmap.items():
            col = _fuzzy_col(col_raw, cols)
            if col not in df2.columns:
                return df, OpResult(name, "mutate", f"update: unknown column {col_raw}",
                                    error=f"Column '{col_raw}' not found")
            _safe_assign(df2, mask, col, _resolve_set_value(val, df2, mask, col))
        sets = ", ".join(f"{_fuzzy_col(c, cols)}={v}" for c, v in setmap.items())
        return df2, OpResult(name, "mutate",
                             f"Set {sets} where {_describe_filter(filt)}", affected=n)

    # ── delete ───────────────────────────────────────────────────────────
    if name in ("delete", "delete_row"):
        filt = op.get("filter")
        if filt is None and op.get("match_col"):
            filt = {"match": "all", "conditions":
                    [{"column": op["match_col"], "op": "eq", "value": op.get("match_val")}]}
        mask = _build_mask(df, filt)
        n = int(mask.sum())
        df2 = df[~mask].reset_index(drop=True)
        return df2, OpResult(name, "mutate",
                             f"Delete {n} row(s) where {_describe_filter(filt)}", affected=n)

    # ── clear_rows ───────────────────────────────────────────────────────
    if name == "clear_rows":
        n = len(df)
        return pd.DataFrame(columns=cols), OpResult(name, "mutate",
                                                    f"Clear all {n} row(s)", affected=n)

    # ── add_column ───────────────────────────────────────────────────────
    if name == "add_column":
        col = str(op.get("name", "")).strip()
        if not col:
            return df, OpResult(name, "mutate", "add_column: no name", error="no name")
        if col in df.columns:
            return df, OpResult(name, "mutate", f"add_column: '{col}' exists",
                                error=f"Column '{col}' already exists")
        df2 = df.copy()
        formula = op.get("formula")
        if formula:
            df2[col] = _eval_formula(formula.lstrip("="), df2)
            summary = f"Add column '{col}' = {formula}"
        else:
            df2[col] = _coerce_value(op.get("default", "")) if op.get("default") not in (None, "") else ""
            summary = f"Add column '{col}'"
        return df2, OpResult(name, "mutate", summary, affected=len(df2))

    # ── delete_column ────────────────────────────────────────────────────
    if name == "delete_column":
        col = _fuzzy_col(str(op.get("name", "")), cols)
        if col not in df.columns:
            return df, OpResult(name, "mutate", f"delete_column: '{op.get('name')}' not found",
                                error=f"Column '{op.get('name')}' not found")
        return df.drop(columns=[col]), OpResult(name, "mutate", f"Delete column '{col}'", affected=1)

    # ── rename_column ────────────────────────────────────────────────────
    if name == "rename_column":
        old = _fuzzy_col(str(op.get("old", "")), cols)
        new = str(op.get("new", "")).strip()
        if old not in df.columns:
            return df, OpResult(name, "mutate", f"rename: '{op.get('old')}' not found",
                                error=f"Column '{op.get('old')}' not found")
        if not new or new in df.columns:
            return df, OpResult(name, "mutate", f"rename: invalid target '{new}'",
                                error=f"Invalid new name '{new}'")
        return df.rename(columns={old: new}), OpResult(name, "mutate",
                                                       f"Rename '{old}' → '{new}'", affected=1)

    # ── sort ─────────────────────────────────────────────────────────────
    if name == "sort":
        by_raw = op.get("by") or op.get("column")
        by = [by_raw] if isinstance(by_raw, str) else list(by_raw or [])
        by = [_fuzzy_col(b, cols) for b in by]
        if not by or any(b not in df.columns for b in by):
            return df, OpResult(name, "mutate", "sort: unknown column", error="unknown sort column")
        asc = op.get("ascending", True)
        # Sort numerically where possible.
        sort_df = df.copy()
        helpers = []
        for b in by:
            num = pd.to_numeric(sort_df[b], errors="coerce")
            if num.notna().any():
                hcol = f"__sort_{b}"
                sort_df[hcol] = num
                helpers.append(hcol)
            else:
                helpers.append(b)
        df2 = sort_df.sort_values(by=helpers, ascending=asc).drop(columns=[h for h in helpers if h.startswith("__sort_")]).reset_index(drop=True)
        order = "↑" if asc else "↓"
        return df2, OpResult(name, "mutate", f"Sort by {', '.join(by)} {order}", affected=len(df2))

    # ── aggregate (read-only) ────────────────────────────────────────────
    if name == "aggregate":
        func = str(op.get("func", "sum")).lower()
        filt = op.get("filter")
        mask = _build_mask(df, filt)
        sub = df[mask]
        group_by = op.get("group_by")
        expr = op.get("expr") or op.get("formula")

        if func == "count" and not op.get("column") and not expr:
            return df, OpResult(name, "read", f"Count where {_describe_filter(filt)}",
                                value=int(mask.sum()))

        # Computed aggregate, e.g. sum of [Qty]*[Price]  ("stock value").
        if expr and not group_by:
            try:
                series = _eval_formula(str(expr).lstrip("="), sub)
            except ValueError as e:
                return df, OpResult(name, "read", "aggregate: bad expr", error=str(e))
            val = _agg_series(series, func)
            return df, OpResult(name, "read",
                                f"{func}({expr}) where {_describe_filter(filt)}",
                                value=_clean_num(val))

        col = _fuzzy_col(str(op.get("column", "")), cols)
        if col not in df.columns:
            col = _infer_numeric_col(df)   # tolerate a missing/loose column
        if col not in df.columns:
            return df, OpResult(name, "read", f"aggregate: unknown column {op.get('column')}",
                                error=f"Column '{op.get('column')}' not found")
        series = pd.to_numeric(sub[col], errors="coerce") if func not in ("count",) else sub[col]

        def _agg(s):
            return {"sum": s.sum(), "mean": s.mean(), "avg": s.mean(), "min": s.min(),
                    "max": s.max(), "median": s.median(), "std": s.std(),
                    "count": s.count()}.get(func, s.sum())

        if group_by:
            gb = _fuzzy_col(group_by, cols)
            grp = sub.copy()
            grp[col] = pd.to_numeric(grp[col], errors="coerce")
            table = grp.groupby(gb)[col].agg(func if func != "avg" else "mean").reset_index()
            table.columns = [gb, f"{func}({col})"]
            return df, OpResult(name, "read", f"{func} of {col} by {gb}", table=table)
        return df, OpResult(name, "read",
                            f"{func}({col}) where {_describe_filter(filt)}",
                            value=_clean_num(_agg_series(series, func)))

    # ── top_n (read-only) ────────────────────────────────────────────────
    if name == "top_n":
        n = int(op.get("n", 5))
        asc = bool(op.get("ascending", False))
        filt = op.get("filter")
        sub = df[_build_mask(df, filt)].copy()
        expr = op.get("expr") or op.get("formula")
        if expr:   # rank by a computed value, e.g. highest [Qty]*[Price]
            try:
                sub["__rank"] = _eval_formula(str(expr).lstrip("="), sub)
            except ValueError as e:
                return df, OpResult(name, "read", "top_n: bad expr", error=str(e))
            rank_label = str(expr)
        else:
            col = _fuzzy_col(str(op.get("column", "")), cols)
            if col not in df.columns:
                col = op.get("by") and _fuzzy_col(str(op["by"]), cols) or _infer_numeric_col(df)
            if col not in df.columns:
                return df, OpResult(name, "read", f"top_n: unknown column {op.get('column')}",
                                    error=f"Column '{op.get('column')}' not found")
            sub["__rank"] = pd.to_numeric(sub[col], errors="coerce")
            rank_label = col
        sub = sub.sort_values("__rank", ascending=asc).drop(columns="__rank").head(n)
        show = op.get("show_columns")
        if show:
            keep = [_fuzzy_col(c, cols) for c in show if _fuzzy_col(c, cols) in df.columns]
            if keep:
                sub = sub[keep]
        word = "bottom" if asc else "top"
        return df, OpResult(name, "read", f"{word.title()} {n} by {rank_label}",
                            table=sub.reset_index(drop=True))

    # ── describe (read-only) ─────────────────────────────────────────────
    if name == "describe":
        return df, OpResult(name, "read", "Summary statistics",
                            table=df.describe(include="all").reset_index())

    return df, OpResult(name or "?", "read", f"Unknown operation '{name}'",
                        error=f"Unknown operation '{name}'")


# ════════════════════════════════════════════════════════════════════════
# Dry-run preview + commit
# ════════════════════════════════════════════════════════════════════════
def dry_run(plan: Plan, df: pd.DataFrame) -> PreviewResult:
    """Apply the plan to a COPY of df, collecting per-step results + a final diff."""
    work = df.copy()
    steps: list[OpResult] = []
    before = df.copy()
    try:
        for op in plan.operations:
            work, res = _exec_op(work, op)
            steps.append(res)
            if res.error:
                # Stop at first hard error so the user sees exactly what failed.
                return PreviewResult(plan, steps, before=before, after=None, error=res.error)
    except Exception as e:  # noqa: BLE001
        return PreviewResult(plan, steps, before=before, after=None, error=str(e))
    return PreviewResult(plan, steps, before=before, after=_sanitize_formulas(work))


def commit(plan: Plan, store) -> PreviewResult:
    """
    Execute the plan's mutations against the live store (with history + save).
    Read-only ops are computed too so the UI can show their results.
    Returns a PreviewResult whose `after` reflects the committed sheet.
    """
    df = store.df
    work = df.copy()
    steps: list[OpResult] = []
    mutated = False
    try:
        for op in plan.operations:
            work, res = _exec_op(work, op)
            steps.append(res)
            if res.error:
                return PreviewResult(plan, steps, before=df, after=None, error=res.error)
            if res.kind == "mutate":
                mutated = True
    except Exception as e:  # noqa: BLE001
        return PreviewResult(plan, steps, before=df, after=None, error=str(e))

    if mutated:
        work = _sanitize_formulas(work)   # never persist a "=[..]" DSL string
        store._push_history()
        store._df = work
        store.save()   # RuntimeError here propagates to the UI (save failures matter)
    return PreviewResult(plan, steps, before=df, after=work)


# ════════════════════════════════════════════════════════════════════════
# Planning  (LLM tool-call → Plan)  with local fast-path fallback
# ════════════════════════════════════════════════════════════════════════
_PLAN_TOOL = {
    "type": "function",
    "function": {
        "name": "build_plan",
        "description": (
            "Translate the user's spreadsheet request into an ordered list of typed "
            "operations. Use read-only operations (aggregate, top_n, describe) to "
            "ANSWER questions, and mutating operations to CHANGE the sheet. Always "
            "reference real column names from the schema."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "summary": {
                    "type": "string",
                    "description": "One short sentence describing what you are doing.",
                },
                "operations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "op": {
                                "type": "string",
                                "enum": ["add_row", "update", "delete", "clear_rows",
                                         "add_column", "delete_column", "rename_column",
                                         "sort", "aggregate", "top_n", "describe"],
                            },
                            "values": {"type": "object", "description": "add_row: {column: value}"},
                            "set": {"type": "object",
                                    "description": "update: {column: value}. Value may be a literal, "
                                                   "a delta like '+5'/'-10', or a formula like "
                                                   "'=[Qty]*[Price]'."},
                            "filter": {
                                "type": "object",
                                "properties": {
                                    "match": {"type": "string", "enum": ["all", "any"]},
                                    "conditions": {
                                        "type": "array",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "column": {"type": "string"},
                                                "op": {"type": "string",
                                                       "enum": ["eq", "ne", "gt", "gte", "lt",
                                                                "lte", "contains", "startswith",
                                                                "endswith", "between", "in",
                                                                "isblank", "notblank"]},
                                                "value": {},
                                            },
                                            "required": ["column", "op"],
                                        },
                                    },
                                },
                            },
                            "name": {"type": "string", "description": "column name for add/delete column"},
                            "formula": {"type": "string", "description": "add_column formula, e.g. '[Qty]*[Price]'"},
                            "default": {"description": "add_column default value"},
                            "old": {"type": "string"},
                            "new": {"type": "string"},
                            "by": {"type": "array", "items": {"type": "string"}, "description": "sort columns"},
                            "ascending": {"description": "bool or list of bools"},
                            "func": {"type": "string",
                                     "enum": ["sum", "mean", "avg", "min", "max", "median", "std", "count"]},
                            "column": {"type": "string"},
                            "expr": {"type": "string",
                                     "description": "aggregate over a formula, e.g. '[Qty]*[Price]' for total value"},
                            "group_by": {"type": "string"},
                            "n": {"type": "integer"},
                            "show_columns": {"type": "array", "items": {"type": "string"}},
                        },
                        "required": ["op"],
                    },
                },
            },
            "required": ["operations"],
        },
    },
}

_SYSTEM = """You are SheetMind, an expert spreadsheet operator. Convert the user's \
request into a precise plan by calling the build_plan function. Follow the schema \
EXACTLY — every operation MUST include all fields shown in the matching example.

Use EXACT column names from the schema. Build filters with the right operator \
(gt, lt, gte, lte, contains, between, in, …); only use eq for exact equality.

REQUIRED FIELDS per op (copy these shapes precisely):
- Increase/decrease/percentage change → update with a `set` map. Use a delta \
("+5"/"-3") for flat changes, or a formula "=[Col]*1.15" for percentages:
  {"op":"update","filter":{"conditions":[{"column":"Category","op":"eq","value":"Writing"}]},"set":{"Price":"=[Price]*1.15"}}
- Total / sum / average / count / min / max → aggregate. For "total value" or any \
sum of a product, use `expr`, NOT a single column:
  {"op":"aggregate","func":"sum","expr":"[Qty]*[Price]"}
  {"op":"aggregate","func":"mean","column":"Price"}
- "top N" / "highest" / "largest" / "which is biggest" → top_n WITH `column` (or \
`expr` to rank by a computed value like stock value):
  {"op":"top_n","column":"Qty","n":3,"ascending":false}
  {"op":"top_n","expr":"[Qty]*[Price]","n":1,"ascending":false}
- Breakdown "by X" (a per-group table) → aggregate with `group_by`. But to restrict \
to ONE subset (e.g. "average price of writing items"), use a `filter`, NOT group_by:
  {"op":"aggregate","func":"sum","column":"Qty","group_by":"Category"}
  {"op":"aggregate","func":"mean","column":"Price","filter":{"conditions":[{"column":"Category","op":"eq","value":"Writing"}]}}
- New computed column → add_column with a formula:
  {"op":"add_column","name":"Total","formula":"[Qty]*[Price]"}
- Add a record → {"op":"add_row","values":{"Item":"Glue","Qty":12}}
- Delete matching rows → {"op":"delete","filter":{"conditions":[{"column":"Qty","op":"lt","value":5}]}}
- Sort → {"op":"sort","by":["Qty"],"ascending":false}
- Rename → {"op":"rename_column","old":"Qty","new":"Quantity"}

Never guess numbers — emit the read op and the engine computes it exactly. \
"Stock value"/"total value" almost always means SUM of quantity × price → use expr.

CRITICAL — do EXACTLY what is asked, nothing more:
- Emit the MINIMUM operations required. Never add, sort, reformat, or "improve" \
anything the user did not explicitly request.
- Do not invent columns, rows, or values. If a needed column does not exist in the \
schema, return empty operations and ask in summary instead of guessing.
- If the request is ambiguous or impossible, return an empty operations list and put \
a clarifying question in summary.
Output ONLY via build_plan."""


def _schema_context(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    dtypes = {c: ("number" if pd.api.types.is_numeric_dtype(df[c]) else "text") for c in cols}
    sample = df.head(3).to_dict(orient="records")
    stats = {"rows": int(len(df)), "columns": cols, "dtypes": dtypes}
    return ("Sheet schema:\n" + json.dumps(stats, default=str)
            + "\nSample rows:\n" + json.dumps(sample, default=str))


def _parse_tool_call(message: dict) -> Optional[dict]:
    calls = message.get("tool_calls") or []
    for c in calls:
        fn = c.get("function", {})
        if fn.get("name") == "build_plan":
            raw = fn.get("arguments") or "{}"
            parsed = None
            try:
                parsed = json.loads(raw)
            except (json.JSONDecodeError, TypeError):
                fb, lb = raw.find("{"), raw.rfind("}")
                if fb != -1 and lb > fb:
                    try:
                        parsed = json.loads(raw[fb:lb + 1])
                    except json.JSONDecodeError:
                        parsed = None
            # Some models double-encode the arguments (a JSON string of JSON).
            if isinstance(parsed, str):
                try:
                    parsed = json.loads(parsed)
                except json.JSONDecodeError:
                    parsed = None
            if isinstance(parsed, dict):
                return parsed
    return None


def interpret(text: str, df: pd.DataFrame) -> Plan:
    """
    Turn a natural-language command into a Plan. Tries a fast local regex
    parser first (zero-latency for common phrasings), then the LLM.
    """
    text = (text or "").strip()
    if not text:
        return Plan(operations=[], summary="Empty command.")
    if len(text) > 1000:
        text = text[:1000]

    local = _local_plan(text, df)
    if local is not None:
        return local

    if not is_configured():
        raise LLMNotConfigured(
            "No OPENROUTER_API_KEY set and the command did not match a built-in "
            "pattern. Add a key from https://openrouter.ai/keys to unlock full AI."
        )

    messages = [
        {"role": "system", "content": _SYSTEM},
        {"role": "user", "content": f"{_schema_context(df)}\n\nRequest: {text}"},
    ]
    return _plan_with_repair(messages, df, max_repairs=2)


_FORCE_PLAN = {"type": "function", "function": {"name": "build_plan"}}


def _validate_plan(plan: "Plan", df: pd.DataFrame) -> list[str]:
    """
    Catch malformed/invalid operations *before* they reach the user by dry-running
    the plan on a copy. Returns a list of human-readable error strings (empty = OK).
    An empty operations list is valid (it means "clarifying question in summary").
    """
    if not plan.operations:
        return []
    pv = dry_run(plan, df)
    errs: list[str] = []
    for step in pv.steps:
        if step.error:
            errs.append(f"operation '{step.op}': {step.error}")
    if pv.error and not errs:
        errs.append(pv.error)
    return errs


def _plan_with_repair(messages: list[dict], df: pd.DataFrame, max_repairs: int = 2) -> "Plan":
    """
    Structured-reflection loop: ask the model for a plan, validate it against the
    real sheet, and on any error feed the precise problem back so it can correct
    itself. This is what lets even small free models hit high accuracy.
    """
    msg = chat(messages, tools=[_PLAN_TOOL], tool_choice=_FORCE_PLAN)
    args = _parse_tool_call(msg)
    if args is None:
        content = (msg.get("content") or "").strip()
        return Plan(operations=[], summary=content or "Could not build a plan.", raw=msg)

    plan = Plan(operations=args.get("operations") or [], summary=args.get("summary", ""), raw=args)

    for _ in range(max_repairs):
        errs = _validate_plan(plan, df)
        if not errs:
            return plan
        call_id = ""
        try:
            call_id = (msg.get("tool_calls") or [{}])[0].get("id", "") or ""
        except (IndexError, AttributeError):
            pass
        messages.append(msg)  # the assistant's tool call
        messages.append({
            "role": "tool",
            "tool_call_id": call_id,
            "name": "build_plan",
            "content": (
                "The plan was REJECTED by the spreadsheet engine:\n- "
                + "\n- ".join(errs)
                + "\n\nCall build_plan again with corrected operations. Remember: "
                  "update needs a `set` map (e.g. {\"Price\":\"=[Price]*1.15\"}); "
                  "aggregate/top_n need `column` or `expr`. Fix only what's broken."
            ),
        })
        msg = chat(messages, tools=[_PLAN_TOOL], tool_choice=_FORCE_PLAN)
        args = _parse_tool_call(msg)
        if args is None:
            break
        plan = Plan(operations=args.get("operations") or [], summary=args.get("summary", ""), raw=args)

    return plan


# ── Local fast-path parser (covers the most common phrasings instantly) ───
def _local_plan(text: str, df: pd.DataFrame) -> Optional[Plan]:
    cols = list(df.columns)
    tl = text.strip().lower()

    # add/increase N to COL where COL is VAL
    m = re.search(r"(add|increase|plus|subtract|decrease|minus)\s+([\d.]+)\s+(?:to|from)\s+"
                  r"(\w[\w\s]*?)\s+where\s+(\w[\w\s]*?)\s+(?:is|=)\s+(.+)", tl)
    if m:
        op, amt = m.group(1), float(m.group(2))
        tc, mc = _fuzzy_col(m.group(3).strip(), cols), _fuzzy_col(m.group(4).strip(), cols)
        mv = m.group(5).strip().strip("\"' ")
        if tc in cols and mc in cols:
            delta = amt if op in ("add", "increase", "plus") else -amt
            d = f"+{delta:g}" if delta >= 0 else f"{delta:g}"
            return Plan(source="local", summary=f"{op} {amt} in {tc}",
                        operations=[{"op": "update",
                                     "filter": {"match": "all", "conditions":
                                                [{"column": mc, "op": "eq", "value": mv}]},
                                     "set": {tc: d}}])
    # set COL to VAL where COL is VAL
    m = re.search(r"(?:set|update|change)\s+(\w[\w\s]*?)\s+to\s+(.+?)\s+where\s+"
                  r"(\w[\w\s]*?)\s+(?:is|=)\s+(.+)", tl)
    if m:
        tc, mc = _fuzzy_col(m.group(1).strip(), cols), _fuzzy_col(m.group(3).strip(), cols)
        val, mv = m.group(2).strip().strip("\"' "), m.group(4).strip().strip("\"' ")
        if tc in cols and mc in cols:
            return Plan(source="local", summary=f"set {tc}",
                        operations=[{"op": "update",
                                     "filter": {"match": "all", "conditions":
                                                [{"column": mc, "op": "eq", "value": mv}]},
                                     "set": {tc: val}}])
    # delete row where COL is VAL
    m = re.search(r"(?:delete|remove)\s+(?:row\s+)?(?:where\s+)?(\w[\w\s]*?)\s+(?:is|=)\s+(.+)", tl)
    if m:
        mc, mv = _fuzzy_col(m.group(1).strip(), cols), m.group(2).strip().strip("\"' ")
        if mc in cols:
            return Plan(source="local", summary=f"delete where {mc}={mv}",
                        operations=[{"op": "delete",
                                     "filter": {"match": "all", "conditions":
                                                [{"column": mc, "op": "eq", "value": mv}]}}])
    return None


# ── One-shot convenience for non-interactive callers (e.g. Excel add-in) ──
def run(text: str, store, confirm: bool = True) -> PreviewResult:
    """
    Interpret + (optionally execute) in one call. When confirm=False the plan is
    committed immediately; when True only a dry-run preview is returned and the
    caller must call commit(preview.plan, store) after the user approves.
    """
    plan = interpret(text, store.df)
    if not plan.operations:
        return PreviewResult(plan, [], error=None)
    if confirm and plan.needs_confirm:
        return dry_run(plan, store.df)
    return commit(plan, store)
