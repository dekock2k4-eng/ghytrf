"""
make_wrapper.py — generate the xlwings entry-point module for a workbook.

The xlwings "Run main" button imports a Python module named after the ACTIVE
workbook, then calls its main(). This creates that module (forwarding to the
SheetMind add-in) so you don't get "No module named '<Workbook>'" or a
SyntaxError for names with spaces/parentheses.

Usage:
    python make_wrapper.py                 # auto-detect the open Excel workbook
    python make_wrapper.py "Inventory.xlsx"
    python make_wrapper.py Sales           # extension optional
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = (
    '"""xlwings entry point for {wb} — forwards to the SheetMind add-in."""\n'
    "import addin\n\n\n"
    "def main():\n"
    "    addin.main()\n"
)


def _active_workbook_name():
    """Best-effort: ask Excel (via xlwings) for the active workbook's name."""
    try:
        import xlwings as xw
        return xw.apps.active.books.active.name
    except Exception:
        return None


def _valid_module(name: str) -> bool:
    # A Python module name: letters/digits/underscore, not starting with a digit.
    return bool(re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name))


def main():
    if len(sys.argv) >= 2:
        raw = sys.argv[1]
    else:
        raw = _active_workbook_name()
        if not raw:
            print("No filename given and couldn't detect an open Excel workbook.")
            print('Usage: python make_wrapper.py "<WorkbookName>.xlsx"')
            sys.exit(1)
        print(f"Detected open workbook: {raw}")

    name = os.path.splitext(os.path.basename(raw.strip()))[0]

    if not _valid_module(name):
        suggestion = re.sub(r"[^A-Za-z0-9_]", "_", name)
        suggestion = re.sub(r"^(\d)", r"wb_\1", suggestion).strip("_") or "Sheet"
        print(f"\n⚠  '{name}' can't be a Python module name (xlwings needs a clean")
        print("   filename — letters, digits, underscores; no spaces/()/leading digit).")
        print(f"   Rename the workbook to e.g.  {suggestion}.xlsx  then re-run this,")
        print(f"   or:  python make_wrapper.py \"{suggestion}.xlsx\"")
        sys.exit(2)

    path = os.path.join(HERE, f"{name}.py")
    with open(path, "w") as f:
        f.write(TEMPLATE.format(wb=f"{name}.xlsx"))
    print(f"Created {path}")
    print(f'Now: quit Excel (Cmd-Q), reopen "{name}.xlsx", and click Open SheetMind.')


if __name__ == "__main__":
    main()
