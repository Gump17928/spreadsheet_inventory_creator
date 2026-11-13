"""Person View sheet: pick a name and see that person's full software list. Read-only except the picker."""
from openpyxl.styles import Alignment, Border, Protection, Side

import refs
from dynamic_names import PEOPLE_LIST
from layout import FIRST_ROW, N_TITLES, Cat, Inv, Mat, Ppl, Sheet
from lookups import catalog_at, person_by_name, text_or_blank
from styles import (
    BRASS,
    BRASS_DARK,
    CRIMSON,
    INK,
    MUTED,
    ROSE_PALE,
    SIZE_BODY,
    SIZE_FIGURE,
    SIZE_HEADING,
    SIZE_SMALL,
    TAB_REPORT,
    applicability_colours,
    clear_when_blank,
    date_flags,
    fill,
    font,
    header_cell,
    page_setup,
    protect_sheet,
    status_colours,
    style_cells,
    title_block,
)
from validation import list_rule
from vocab import Status

PICKER = "B4"
PICKER_MERGE = "B4:D4"
SUMMARY_LABEL_ROW, SUMMARY_VALUE_ROW = 6, 7
DETAIL_TITLE_ROW, DETAIL_HEADER_ROW, DETAIL_FIRST = 9, 10, 11
DETAIL_LAST = DETAIL_FIRST + N_TITLES - 1
PICKER_ROW_HEIGHT = 26
SUMMARY_VALUE_HEIGHT = 34
DATE_FORMAT = "yyyy-mm-dd"


class PV:
    """Details-table columns."""
    SOFTWARE, GROUP, APPLICABILITY, STATUS, VERSION, APPROVED, ASSIGNMENT, EXPIRY, VERIFIED, NOTES = "ABCDEFGHIJ"


DETAIL_COLUMNS = [
    (PV.SOFTWARE, "Software", 26), (PV.GROUP, "Group / Suite", 18), (PV.APPLICABILITY, "Applicability", 22),
    (PV.STATUS, "Status", 17), (PV.VERSION, "Installed Version", 14), (PV.APPROVED, "Approved Version", 14),
    (PV.ASSIGNMENT, "License Assignment", 22), (PV.EXPIRY, "License Expiry", 16), (PV.VERIFIED, "Last Verified", 15),
    (PV.NOTES, "Notes", 40),
]
TILE_COLUMNS = ("A", "B", "C", "D")
PICKER_COLUMNS = ("B", "C", "D")


def _person_row():
    """Position of the picked person within the Matrix."""
    return f"MATCH({PICKER},{refs.matrix_col(Mat.PERSON)},0)"


def _matrix_value(col):
    return f'IF({PICKER}="","",IFERROR(INDEX({refs.matrix_col(col)},{_person_row()}),""))'


def _write_picker(ws):
    ws["A4"] = "Person ▸"
    ws["A4"].font = font(SIZE_BODY, True, BRASS_DARK)
    ws["A4"].alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(PICKER_MERGE)
    edge = Side(style="medium", color=BRASS)
    for letter in PICKER_COLUMNS:
        ws[f"{letter}4"].fill = fill(ROSE_PALE)
        ws[f"{letter}4"].border = Border(top=edge, bottom=edge, left=edge if letter == PICKER_COLUMNS[0] else None,
                                         right=edge if letter == PICKER_COLUMNS[-1] else None)
    ws[PICKER].font = font(SIZE_HEADING, True)
    ws[PICKER].alignment = Alignment(vertical="center")
    ws[PICKER].protection = Protection(locked=False)
    ws.row_dimensions[4].height = PICKER_ROW_HEIGHT
    list_rule(ws, PICKER, f"={PEOPLE_LIST}", "Person", "Choose a person from the dropdown.")


def _write_person_facts(ws):
    """Role / department / computer of the picked person, to the right of the picker."""
    facts = [("Role", Ppl.ROLE), ("Dept", Ppl.DEPARTMENT), ("PC", Ppl.COMPUTER)]
    joined = ' & "   |   " & '.join(f'"{label}: "&{person_by_name(col, PICKER)}' for label, col in facts)
    ws["E4"] = f'=IF({PICKER}="","",IFERROR({joined},""))'
    ws["E4"].font = font(SIZE_SMALL, color=MUTED, italic=True)
    ws["E4"].alignment = Alignment(vertical="center")


def _write_figures(ws):
    """Four flat figures: label in brass, number in black (crimson for Missing), a brass rule above."""
    detail_status = f"${PV.STATUS}${DETAIL_FIRST}:${PV.STATUS}${DETAIL_LAST}"
    not_applicable = f'IF({PICKER}="","",COUNTIF({detail_status},"{Status.NOT_APPLICABLE}"))'
    figures = [
        ("HAVE", _matrix_value(Mat.HAVE), INK),
        ("MISSING", _matrix_value(Mat.MISSING), CRIMSON),
        ("NEEDS ATTENTION", _matrix_value(Mat.ATTENTION), INK),
        ("NOT APPLICABLE", not_applicable, MUTED),
    ]
    rule = Border(top=Side(style="medium", color=BRASS))
    for letter, (label, formula, colour) in zip(TILE_COLUMNS, figures, strict=True):
        ws[f"{letter}{SUMMARY_LABEL_ROW}"] = label
        ws[f"{letter}{SUMMARY_VALUE_ROW}"] = "=" + formula
        ws[f"{letter}{SUMMARY_LABEL_ROW}"].font = font(SIZE_SMALL, True, BRASS_DARK)
        ws[f"{letter}{SUMMARY_VALUE_ROW}"].font = font(SIZE_FIGURE, True, colour)
        ws[f"{letter}{SUMMARY_LABEL_ROW}"].border = rule
        for row in (SUMMARY_LABEL_ROW, SUMMARY_VALUE_ROW):
            ws[f"{letter}{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[SUMMARY_VALUE_ROW].height = SUMMARY_VALUE_HEIGHT


def _detail_formulas(row, slot):
    """Formulas for one title's row in the details table."""
    name = f"${PV.SOFTWARE}{row}"
    offset = f"(MATCH({PICKER},{refs.ppl(Ppl.NAME)},0)-1)*{N_TITLES}+{slot + 1}"   # that person's Inventory row
    catalog_row = FIRST_ROW + slot

    def from_inventory(col):
        return f"INDEX({refs.inv(col)},{offset})"

    def date_from_inventory(col):
        return f'IFERROR(IF({from_inventory(col)}="","",{from_inventory(col)}),"")'

    body = {
        PV.GROUP: text_or_blank(catalog_at(Cat.GROUP, catalog_row)),
        PV.APPLICABILITY: text_or_blank(from_inventory(Inv.APPLICABILITY)),
        PV.STATUS: text_or_blank(from_inventory(Inv.STATUS)),
        PV.VERSION: text_or_blank(from_inventory(Inv.VERSION)),
        PV.APPROVED: text_or_blank(catalog_at(Cat.APPROVED, catalog_row)),
        PV.ASSIGNMENT: text_or_blank(from_inventory(Inv.ASSIGNMENT)),
        PV.EXPIRY: date_from_inventory(Inv.EXPIRY),
        PV.VERIFIED: date_from_inventory(Inv.VERIFIED_ON),
        PV.NOTES: text_or_blank(from_inventory(Inv.NOTES)),
    }
    formulas = {letter: f'=IF({name}="","",{expr})' for letter, expr in body.items()}
    catalog_name = catalog_at(Cat.NAME, catalog_row)
    formulas[PV.SOFTWARE] = f'=IF(OR({PICKER}="",{catalog_name}=""),"",{catalog_name})'
    return formulas


def _write_details(ws):
    ws[f"A{DETAIL_TITLE_ROW}"] = "DETAILS"
    ws[f"A{DETAIL_TITLE_ROW}"].font = font(SIZE_BODY, True, BRASS_DARK)
    for index, (_letter, title, width) in enumerate(DETAIL_COLUMNS, start=1):
        header_cell(ws, DETAIL_HEADER_ROW, index, title, width=width)
    for slot in range(N_TITLES):
        row = DETAIL_FIRST + slot
        for letter, formula in _detail_formulas(row, slot).items():
            ws[f"{letter}{row}"] = formula
    style_cells(ws, f"{PV.SOFTWARE}{DETAIL_FIRST}:{PV.NOTES}{DETAIL_LAST}")
    style_cells(ws, f"{PV.APPLICABILITY}{DETAIL_FIRST}:{PV.VERIFIED}{DETAIL_LAST}", align="center")
    for letter in (PV.EXPIRY, PV.VERIFIED):
        style_cells(ws, f"{letter}{DETAIL_FIRST}:{letter}{DETAIL_LAST}", fmt=DATE_FORMAT, align="center")


def _add_formatting(ws):
    status_colours(ws, f"{PV.STATUS}{DETAIL_FIRST}:{PV.STATUS}{DETAIL_LAST}")
    applicability_colours(ws, f"{PV.APPLICABILITY}{DETAIL_FIRST}:{PV.APPLICABILITY}{DETAIL_LAST}")
    date_flags(ws, f"{PV.EXPIRY}{DETAIL_FIRST}:{PV.EXPIRY}{DETAIL_LAST}")
    clear_when_blank(ws, f"{PV.SOFTWARE}{DETAIL_FIRST}:{PV.NOTES}{DETAIL_LAST}", f"${PV.SOFTWARE}{DETAIL_FIRST}")


def create(wb):
    ws = wb.create_sheet(Sheet.PERSON_VIEW)
    title_block(ws, "Person View", "Pick a name to see what that person has, is missing, or doesn't need.")
    ws.sheet_properties.tabColor = TAB_REPORT
    _write_picker(ws)
    _write_person_facts(ws)
    _write_figures(ws)
    _write_details(ws)
    _add_formatting(ws)
    ws.freeze_panes = f"A{DETAIL_FIRST}"
    ws.auto_filter.ref = f"A{DETAIL_HEADER_ROW}:{PV.NOTES}{DETAIL_LAST}"
    page_setup(ws, f"{DETAIL_HEADER_ROW}:{DETAIL_HEADER_ROW}")
    protect_sheet(ws)


def select_person(ws, name):
    """Pre-select someone (used by the example workbook)."""
    ws[PICKER] = name
