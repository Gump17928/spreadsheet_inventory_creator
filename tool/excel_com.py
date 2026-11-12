"""Drives Excel over COM (pywin32) to open a workbook, let it calculate, and read values back.

Dev-only: the finished workbooks do not need this. Excel must be installed.
"""
from contextlib import contextmanager
from pathlib import Path

XL_FORMULAS = -4123    # SpecialCells type: cells that contain formulas
XL_ERRORS = 16         # SpecialCells value: formulas whose result is an error
NO_CELLS_HRESULT = -2146827284   # SpecialCells raises this when nothing matches


def excel_available():
    """True when pywin32 is installed and Excel is registered on this PC."""
    try:
        import winreg

        import win32com.client  # noqa: F401  (imported only to prove it is installed)
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
    except (ImportError, OSError):
        return False
    return True


@contextmanager
def excel_session():
    """A hidden Excel instance that is always quit, even if the caller fails."""
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        app.Quit()


@contextmanager
def open_workbook(app, path):
    """Open a workbook, force a full recalculation, and close it without saving."""
    workbook = app.Workbooks.Open(str(Path(path).resolve()))   # Excel does not understand paths relative to our cwd
    try:
        app.CalculateFullRebuild()
        yield workbook
    finally:
        workbook.Close(SaveChanges=False)


def error_cells(workbook):
    """Addresses of every formula cell that evaluates to an error, as {sheet name: [address, ...]}."""
    import pywintypes
    found = {}
    for sheet in workbook.Worksheets:
        try:
            cells = sheet.UsedRange.SpecialCells(XL_FORMULAS, XL_ERRORS)
        except pywintypes.com_error as error:
            if error.hresult != NO_CELLS_HRESULT and error.excepinfo[5] != NO_CELLS_HRESULT:
                raise
            continue
        found[sheet.Name] = [cell.Address.replace("$", "") for cell in cells]
    return found


def read_cell(workbook, sheet, address):
    return workbook.Worksheets(sheet).Range(address).Value


def read_row(workbook, sheet, row, first_col, last_col):
    """Values of one row between two column letters, as a list."""
    values = workbook.Worksheets(sheet).Range(f"{first_col}{row}:{last_col}{row}").Value
    return list(values[0])
