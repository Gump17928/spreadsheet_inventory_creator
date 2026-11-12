"""Opens a workbook in Excel, recalculates it and reports formula errors plus a status summary.

Usage:  .venv\\Scripts\\python.exe tool\\verify_in_excel.py [workbook.xlsx]
"""
import sys
from collections import Counter
from pathlib import Path

import excel_com
from layout import FIRST_ROW, INV_LAST, MATRIX_FIRST_TITLE, MATRIX_LAST_TITLE, PEOPLE_LAST, Inv, Mat, Sheet

DEFAULT_WORKBOOK = Path(__file__).resolve().parent.parent / "Software_Inventory_Example.xlsx"


def status_counts(workbook, sheet, first_col, last_col, first_row, last_row):
    """How often each non-blank value appears in a block of cells."""
    block = workbook.Worksheets(sheet).Range(f"{first_col}{first_row}:{last_col}{last_row}").Value
    return Counter(value for row in block for value in row if value not in (None, ""))


def main(path):
    with excel_com.excel_session() as app, excel_com.open_workbook(app, path) as workbook:
        errors = excel_com.error_cells(workbook)
        print("Formula errors:", errors or "none")
        inventory = status_counts(workbook, Sheet.INVENTORY, Inv.STATUS, Inv.STATUS, FIRST_ROW, INV_LAST)
        matrix = status_counts(workbook, Sheet.MATRIX, MATRIX_FIRST_TITLE, MATRIX_LAST_TITLE, FIRST_ROW, PEOPLE_LAST)
        print("Inventory statuses:", dict(inventory))
        print("Matrix statuses:   ", dict(matrix))
        people = status_counts(workbook, Sheet.MATRIX, Mat.PERSON, Mat.PERSON, FIRST_ROW, PEOPLE_LAST)
        print("Matrix people:", people.total())
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WORKBOOK))
