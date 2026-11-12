"""Absolute range references to the data areas of each sheet, e.g. `Inventory!$S$6:$S$1505`."""
from layout import (
    CATALOG_LAST,
    FIRST_ROW,
    HEADER_ROW,
    INV_LAST,
    MATRIX_FIRST_TITLE,
    MATRIX_LAST_TITLE,
    PEOPLE_LAST,
    ROLE_COL,
    ROLE_GRID_LAST_COL,
    ROLES_LAST,
    Sheet,
)


def quote(sheet):
    """Sheet name as it must appear in a formula."""
    return f"'{sheet}'" if " " in sheet else sheet


def column(sheet, col, last_row):
    """One column of data rows: `Sheet!$C$6:$C$99`."""
    return f"{quote(sheet)}!${col}${FIRST_ROW}:${col}${last_row}"


def inv(col):
    return column(Sheet.INVENTORY, col, INV_LAST)


def cat(col):
    return column(Sheet.CATALOG, col, CATALOG_LAST)


def ppl(col):
    return column(Sheet.PEOPLE, col, PEOPLE_LAST)


def matrix_col(col):
    return column(Sheet.MATRIX, col, PEOPLE_LAST)


def matrix_body():
    """The whole person x title block of the Matrix."""
    return f"{quote(Sheet.MATRIX)}!${MATRIX_FIRST_TITLE}${FIRST_ROW}:${MATRIX_LAST_TITLE}${PEOPLE_LAST}"


def growing_list(sheet, col, last_row):
    """A range that covers only the filled cells of a column, so dropdowns never show blanks."""
    first = f"{quote(sheet)}!${col}${FIRST_ROW}"
    return f"OFFSET({first},0,0,MAX(1,COUNTA({column(sheet, col, last_row)})),1)"


ROLE_NAMES = column(Sheet.ROLES, ROLE_COL, ROLES_LAST)
ROLE_GRID = f"{quote(Sheet.ROLES)}!$B${FIRST_ROW}:${ROLE_GRID_LAST_COL}${ROLES_LAST}"
ROLE_TITLES = f"{quote(Sheet.ROLES)}!$B${HEADER_ROW}:${ROLE_GRID_LAST_COL}${HEADER_ROW}"
