"""
xlwings entry point for the workbook **Test.xlsx**.

On macOS the xlwings "Run main" button imports a Python module named after the
ACTIVE WORKBOOK and calls its main(). So a workbook called `Test.xlsx` needs a
`Test.py` on the PYTHONPATH. This file just forwards to the SheetMind add-in.

For a different workbook (e.g. Inventory.xlsx) create an Inventory.py with the
same two lines, or run:  python make_wrapper.py "Inventory.xlsx"
"""
import addin


def main():
    addin.main()
