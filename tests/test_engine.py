"""
Deterministic tests for the SheetMind engine.

These cover the parts that must never silently break: condition evaluation,
formula math, every operation's executor, dry-run vs commit, and the local
fast-path parser. The LLM is NOT exercised here — planning is tested by feeding
the executor pre-built plans (the same dict shape the LLM emits).
"""
import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import engine
from engine import Plan, dry_run, commit, _build_mask, _eval_formula, _exec_op, _local_plan


@pytest.fixture
def df():
    return pd.DataFrame({
        "Item": ["Pen", "Pencil", "Eraser", "Marker"],
        "Qty": [10, 50, 5, 20],
        "Price": [2.0, 1.0, 3.0, 8.0],
        "Category": ["Writing", "Writing", "Office", "Writing"],
    })


class FakeStore:
    """Minimal store implementing the interface engine.commit() needs."""
    def __init__(self, df):
        self._df = df
        self.saved = 0
        self.history = []

    @property
    def df(self):
        return self._df

    def _push_history(self):
        self.history.append(self._df.copy())

    def save(self):
        self.saved += 1


# ── conditions ────────────────────────────────────────────────────────────
def test_mask_numeric_gt(df):
    mask = _build_mask(df, {"conditions": [{"column": "Qty", "op": "gt", "value": 10}]})
    assert df[mask]["Item"].tolist() == ["Pencil", "Marker"]


def test_mask_between(df):
    mask = _build_mask(df, {"conditions": [{"column": "Qty", "op": "between", "value": [5, 20]}]})
    assert sorted(df[mask]["Item"].tolist()) == ["Eraser", "Marker", "Pen"]


def test_mask_contains_case_insensitive(df):
    mask = _build_mask(df, {"conditions": [{"column": "Item", "op": "contains", "value": "en"}]})
    assert sorted(df[mask]["Item"].tolist()) == ["Pen", "Pencil"]


def test_mask_any_match(df):
    flt = {"match": "any", "conditions": [
        {"column": "Category", "op": "eq", "value": "Office"},
        {"column": "Qty", "op": "gt", "value": 40},
    ]}
    assert sorted(df[_build_mask(df, flt)]["Item"].tolist()) == ["Eraser", "Pencil"]


def test_empty_filter_is_all(df):
    assert _build_mask(df, None).all()


# ── formulas ──────────────────────────────────────────────────────────────
def test_formula_product(df):
    res = _eval_formula("[Qty] * [Price]", df)
    assert res.tolist() == [20.0, 50.0, 15.0, 160.0]


def test_formula_percentage(df):
    res = _eval_formula("[Price] * 1.1", df)
    assert res.round(2).tolist() == [2.2, 1.1, 3.3, 8.8]


def test_formula_rejects_unsafe(df):
    with pytest.raises(ValueError):
        _eval_formula("__import__('os')", df)


# ── operation executors ───────────────────────────────────────────────────
def test_add_row(df):
    df2, res = _exec_op(df, {"op": "add_row", "values": {"Item": "Glue", "Qty": "three"}})
    assert len(df2) == 5
    assert df2.iloc[-1]["Qty"] == 3          # word-number coercion
    assert res.affected == 1


def test_update_delta(df):
    df2, res = _exec_op(df, {"op": "update",
                             "filter": {"conditions": [{"column": "Item", "op": "eq", "value": "Pen"}]},
                             "set": {"Qty": "+5"}})
    assert df2.loc[df2["Item"] == "Pen", "Qty"].iloc[0] == 15
    assert res.affected == 1


def test_update_formula(df):
    df2, _ = _exec_op(df, {"op": "update",
                           "filter": {"conditions": [{"column": "Category", "op": "eq", "value": "Writing"}]},
                           "set": {"Price": "=[Price]*2"}})
    assert df2.loc[df2["Item"] == "Marker", "Price"].iloc[0] == 16.0
    assert df2.loc[df2["Item"] == "Eraser", "Price"].iloc[0] == 3.0   # untouched


def test_delete_filter(df):
    df2, res = _exec_op(df, {"op": "delete",
                             "filter": {"conditions": [{"column": "Qty", "op": "lt", "value": 10}]}})
    assert "Eraser" not in df2["Item"].tolist()
    assert res.affected == 1


def test_add_column_formula(df):
    df2, _ = _exec_op(df, {"op": "add_column", "name": "Total", "formula": "[Qty]*[Price]"})
    assert df2["Total"].tolist() == [20.0, 50.0, 15.0, 160.0]


def test_sort_numeric_desc(df):
    df2, _ = _exec_op(df, {"op": "sort", "by": ["Qty"], "ascending": False})
    assert df2["Item"].tolist() == ["Pencil", "Marker", "Pen", "Eraser"]


def test_rename_column(df):
    df2, _ = _exec_op(df, {"op": "rename_column", "old": "Qty", "new": "Quantity"})
    assert "Quantity" in df2.columns and "Qty" not in df2.columns


# ── read-only analytics ───────────────────────────────────────────────────
def test_aggregate_sum(df):
    _, res = _exec_op(df, {"op": "aggregate", "func": "sum", "column": "Qty"})
    assert res.value == 85


def test_aggregate_sum_filtered(df):
    _, res = _exec_op(df, {"op": "aggregate", "func": "sum", "column": "Qty",
                           "filter": {"conditions": [{"column": "Category", "op": "eq", "value": "Writing"}]}})
    assert res.value == 80


def test_aggregate_group_by(df):
    _, res = _exec_op(df, {"op": "aggregate", "func": "sum", "column": "Qty", "group_by": "Category"})
    by_cat = dict(zip(res.table["Category"], res.table["sum(Qty)"]))
    assert by_cat["Writing"] == 80 and by_cat["Office"] == 5


def test_top_n(df):
    _, res = _exec_op(df, {"op": "top_n", "column": "Price", "n": 2})
    assert res.table["Item"].tolist() == ["Marker", "Eraser"]


def test_aggregate_expr_stock_value(df):
    # sum of Qty*Price = 20 + 50 + 15 + 160 = 245
    _, res = _exec_op(df, {"op": "aggregate", "func": "sum", "expr": "[Qty]*[Price]"})
    assert res.value == 245


def test_top_n_expr_highest_value(df):
    # highest Qty*Price is Marker (160)
    _, res = _exec_op(df, {"op": "top_n", "expr": "[Qty]*[Price]", "n": 1})
    assert res.table["Item"].tolist() == ["Marker"]


def test_aggregate_count_only(df):
    _, res = _exec_op(df, {"op": "aggregate", "func": "count",
                           "filter": {"conditions": [{"column": "Category", "op": "eq", "value": "Writing"}]}})
    assert res.value == 3


# ── plan-level dry_run / commit ───────────────────────────────────────────
def test_dry_run_does_not_mutate(df):
    plan = Plan(operations=[{"op": "delete", "filter": {"conditions":
                [{"column": "Item", "op": "eq", "value": "Pen"}]}}])
    pv = dry_run(plan, df)
    assert pv.ok
    assert len(df) == 4                       # original untouched
    assert len(pv.after) == 3                 # preview reflects deletion


def test_commit_mutates_store(df):
    store = FakeStore(df.copy())
    plan = Plan(operations=[{"op": "update",
                "filter": {"conditions": [{"column": "Item", "op": "eq", "value": "Pen"}]},
                "set": {"Qty": 99}}])
    pv = commit(plan, store)
    assert pv.ok
    assert store.saved == 1
    assert store.df.loc[store.df["Item"] == "Pen", "Qty"].iloc[0] == 99


def test_commit_readonly_does_not_save(df):
    store = FakeStore(df.copy())
    plan = Plan(operations=[{"op": "aggregate", "func": "sum", "column": "Qty"}])
    pv = commit(plan, store)
    assert store.saved == 0
    assert pv.steps[0].value == 85


def test_bad_column_reports_error(df):
    plan = Plan(operations=[{"op": "delete", "filter": {"conditions":
                [{"column": "Nope", "op": "eq", "value": "x"}]}}])
    pv = dry_run(plan, df)
    assert not pv.ok and pv.error


# ── local fast-path parser ────────────────────────────────────────────────
def test_local_add_delta(df):
    plan = _local_plan("add 50 to Qty where Item is Pen", df)
    assert plan and plan.operations[0]["set"]["Qty"] == "+50"


def test_local_delete(df):
    plan = _local_plan("delete row where Item is Pencil", df)
    assert plan and plan.operations[0]["op"] == "delete"


def test_local_returns_none_for_complex(df):
    assert _local_plan("what is the total quantity of writing items", df) is None
