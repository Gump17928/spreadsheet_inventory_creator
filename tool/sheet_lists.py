"""Lists sheet: dropdown sources (some fixed, some editable) and the two warning thresholds."""
from dataclasses import dataclass

from openpyxl.workbook.defined_name import DefinedName

import refs
import starter_data
import validation
from layout import FIRST_ROW, LIST_LAST, N_LIST_SLOTS, STALE_DAYS, WARN_DAYS, Lst, Sheet, col_index
from styles import TAB_REFERENCE, header_cell, note, page_setup, style_cells, title_block
from vocab import Applicability, InstallState, LicenseAssign, LicenseModel, values

LIST_WIDTH = 26
SETTING_LABEL_WIDTH = 34
SETTING_VALUE_WIDTH = 12


@dataclass(frozen=True)
class ListColumn:
    letter: str
    title: str
    items: list[str]
    editable: bool
    defined_name: str


@dataclass(frozen=True)
class Setting:
    label: str
    cell: str
    default: int
    defined_name: str


LIST_COLUMNS = [
    ListColumn(Lst.INSTALL_STATE, "Install State", values(InstallState), False, "InstallStates"),
    ListColumn(Lst.APPLICABILITY, "Applicability", values(Applicability), False, "ApplicabilityList"),
    ListColumn(Lst.LICENSE_ASSIGN, "License Assignment", values(LicenseAssign), False, "LicenseAssignList"),
    ListColumn(Lst.LICENSE_MODEL, "License Model", values(LicenseModel), True, "LicenseModelList"),
    ListColumn(Lst.GROUP, "Group / Suite", starter_data.GROUPS, True, "GroupList"),
    ListColumn(Lst.CATEGORY, "Category", starter_data.CATEGORIES, True, "CategoryList"),
    ListColumn(Lst.DEPARTMENT, "Department", starter_data.DEPARTMENTS, True, "DeptList"),
]
SETTINGS = [
    Setting("License warning window (days)", Lst.WARN_CELL, 60, WARN_DAYS),
    Setting("Data is stale after (days since verified)", Lst.STALE_CELL, 180, STALE_DAYS),
]


def _absolute(cell):
    """'J6' -> '$J$6'."""
    letters = cell.rstrip("0123456789")
    return f"${letters}${cell[len(letters):]}"


def _write_list(ws, column):
    """Header plus items; editable lists get spare empty slots."""
    index = col_index(column.letter)
    header_cell(ws, 5, index, column.title, calc=not column.editable, width=LIST_WIDTH)
    rows = N_LIST_SLOTS if column.editable else len(column.items)
    for offset, item in enumerate(column.items):
        ws.cell(row=FIRST_ROW + offset, column=index, value=item)
    style_cells(ws, f"{column.letter}{FIRST_ROW}:{column.letter}{FIRST_ROW + rows - 1}", calc=not column.editable)


def _write_settings(ws):
    index = col_index(Lst.SETTING_LABEL)
    header_cell(ws, 5, index, "Setting", width=SETTING_LABEL_WIDTH)
    header_cell(ws, 5, index + 1, "Value", width=SETTING_VALUE_WIDTH)
    for setting in SETTINGS:
        row = int(setting.cell[1:])
        ws[f"{Lst.SETTING_LABEL}{row}"] = setting.label
        ws[setting.cell] = setting.default
        style_cells(ws, f"{Lst.SETTING_LABEL}{row}:{Lst.SETTING_LABEL}{row}")
        style_cells(ws, f"{setting.cell}:{setting.cell}", align="center")
        validation.days_rule(ws, setting.cell)
    note(ws, f"{Lst.SETTING_LABEL}9",
         "Tip: to add a group, category or department, type it in the next empty cell of its column.")


def _list_reference(column):
    """Formula behind a list's defined name; editable lists grow as items are added."""
    if column.editable:
        return refs.growing_list(Sheet.LISTS, column.letter, LIST_LAST)
    return f"{Sheet.LISTS}!${column.letter}${FIRST_ROW}:${column.letter}${FIRST_ROW + len(column.items) - 1}"


def _define_names(wb):
    for setting in SETTINGS:
        wb.defined_names[setting.defined_name] = DefinedName(
            setting.defined_name, attr_text=f"{Sheet.LISTS}!{_absolute(setting.cell)}")
    for column in LIST_COLUMNS:
        wb.defined_names[column.defined_name] = DefinedName(column.defined_name, attr_text=_list_reference(column))


def create(wb):
    ws = wb.create_sheet(Sheet.LISTS)
    title_block(ws, "Lists & Settings", "Dropdown sources and warning thresholds. Grey lists are fixed "
                                        "(formulas depend on them); white lists are yours to edit.")
    ws.sheet_properties.tabColor = TAB_REFERENCE
    for column in LIST_COLUMNS:
        _write_list(ws, column)
    _write_settings(ws)
    _define_names(wb)
    ws.freeze_panes = f"A{FIRST_ROW}"
    page_setup(ws)
