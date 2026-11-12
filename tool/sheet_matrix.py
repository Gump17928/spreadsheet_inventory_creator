"""Matrix sheet: every person against every title, colour-coded. Read-only; all cells are formulas."""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, Side

import refs
from layout import (
    FIRST_ROW,
    HEADER_ROW,
    MATRIX_FIRST_TITLE,
    MATRIX_LAST_TITLE,
    N_TITLES,
    PEOPLE_LAST,
    Cat,
    Inv,
    Mat,
    Ppl,
    Sheet,
    col_index,
)
from lookups import not_installed_status, role_default_by_slot
from styles import (
    BORDER,
    NAVY,
    ORANGE_TEXT,
    TAB_REPORT,
    Col,
    clear_when_blank,
    fill,
    font,
    header_cell,
    page_setup,
    protect_report,
    status_colours,
    style_cells,
    title_block,
    write_headers,
)
from vocab import ATTENTION, HAVE, Status

GROUP_ROW_FILL = "E3EAF4"
GROUP_ROW_HEIGHT = 30
TITLE_HEADER_HEIGHT = 58
TITLE_COL_WIDTH = 13
GROUP_EDGE = Side(style="medium", color=NAVY)
THIN_EDGE = BORDER.left
LEFT_COLS = [
    Col(Mat.PERSON, "Person", 22), Col(Mat.DEPARTMENT, "Department", 15), Col(Mat.ROLE, "Role", 20),
    Col(Mat.COMPUTER, "Computer", 15), Col(Mat.HAVE, "Have", 7, True, align="center"),
    Col(Mat.MISSING, "Missing", 8, True, align="center"),
    Col(Mat.ATTENTION, "Needs Attention", 10, True, align="center"),
]
BODY = f"{MATRIX_FIRST_TITLE}{FIRST_ROW}:{MATRIX_LAST_TITLE}{PEOPLE_LAST}"


def _count(statuses, row_range):
    return "+".join(f'COUNTIF({row_range},"{status}")' for status in statuses)


def _person_formulas(row):
    """Person, department, role, computer and the three counters for one Matrix row."""
    person = f"${Mat.PERSON}{row}"
    row_range = f"${MATRIX_FIRST_TITLE}{row}:${MATRIX_LAST_TITLE}{row}"
    return {
        Mat.PERSON: f'=IF(People!${Ppl.NAME}{row}="","",People!${Ppl.NAME}{row})',
        Mat.DEPARTMENT: f'=IF({person}="","",People!${Ppl.DEPARTMENT}{row}&"")',
        Mat.ROLE: f'=IF({person}="","",People!${Ppl.ROLE}{row}&"")',
        Mat.COMPUTER: f'=IF({person}="","",People!${Ppl.COMPUTER}{row}&"")',
        Mat.HAVE: f'=IF({person}="","",{_count(HAVE, row_range)})',
        Mat.MISSING: f'=IF({person}="","",COUNTIF({row_range},"{Status.MISSING}"))',
        Mat.ATTENTION: f'=IF({person}="","",{_count(ATTENTION, row_range)})',
    }


def cell_formula(row, slot):
    """Status of one person for one title: the Inventory row's status if there is one, else from the role default."""
    letter = Mat.title_col(slot)
    person, title = f"${Mat.PERSON}{row}", f"{letter}${HEADER_ROW}"
    fallback = not_installed_status(role_default_by_slot(f"${Mat.ROLE}{row}", slot))
    lookup = f'INDEX({refs.inv(Inv.STATUS)},MATCH({person}&"|"&{title},{refs.inv(Inv.KEY)},0))'
    return f'=IF(OR({person}="",{title}=""),"",IFERROR({lookup},{fallback}))'


def _catalog_mirror(column, slot):
    """Formula that shows a Catalog value for slot `slot`, or blank when the slot is empty."""
    name = f"{Sheet.CATALOG}!${Cat.NAME}${FIRST_ROW + slot}"
    value = f"{Sheet.CATALOG}!${column}${FIRST_ROW + slot}"
    return f'=IF({name}="","",{value})'


def _write_title_headers(ws):
    """Group names above, title names below; both mirror the Catalog order."""
    for slot in range(N_TITLES):
        index = col_index(Mat.title_col(slot))
        group = ws.cell(row=Mat.GROUP_ROW, column=index, value=_catalog_mirror(Cat.GROUP, slot))
        group.font = font(9, True, NAVY)
        group.fill = fill(GROUP_ROW_FILL)
        group.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_cell(ws, HEADER_ROW, index, _catalog_mirror(Cat.NAME, slot), width=TITLE_COL_WIDTH)
    ws.row_dimensions[Mat.GROUP_ROW].height = GROUP_ROW_HEIGHT
    ws.row_dimensions[HEADER_ROW].height = TITLE_HEADER_HEIGHT


def _write_body(ws):
    for row in range(FIRST_ROW, PEOPLE_LAST + 1):
        for letter, formula in _person_formulas(row).items():
            ws[f"{letter}{row}"] = formula
        for slot in range(N_TITLES):
            ws[f"{Mat.title_col(slot)}{row}"] = cell_formula(row, slot)
    style_cells(ws, f"{Mat.PERSON}{FIRST_ROW}:{Mat.COMPUTER}{PEOPLE_LAST}")
    style_cells(ws, f"{Mat.HAVE}{FIRST_ROW}:{Mat.ATTENTION}{PEOPLE_LAST}", calc=True, align="center")
    for row in ws[BODY]:
        for cell in row:
            cell.font = font(9)
            cell.alignment = Alignment(horizontal="center", vertical="center")


def _add_formatting(ws):
    """Status colours; borders and group dividers appear only where there is data."""
    status_colours(ws, BODY)
    first = MATRIX_FIRST_TITLE
    used = f'AND(${Mat.PERSON}{FIRST_ROW}<>"",{first}${HEADER_ROW}<>"")'
    ws.conditional_formatting.add(BODY, FormulaRule(
        formula=[f'AND({used},{first}${Mat.GROUP_ROW}<>{Mat.ATTENTION}${Mat.GROUP_ROW})'],
        border=Border(left=GROUP_EDGE, right=THIN_EDGE, top=THIN_EDGE, bottom=THIN_EDGE)))
    ws.conditional_formatting.add(BODY, FormulaRule(
        formula=[used], border=Border(left=THIN_EDGE, right=THIN_EDGE, top=THIN_EDGE, bottom=THIN_EDGE)))
    clear_when_blank(ws, f"{Mat.PERSON}{FIRST_ROW}:{Mat.ATTENTION}{PEOPLE_LAST}", f"${Mat.PERSON}{FIRST_ROW}")
    header_cells = f"{MATRIX_FIRST_TITLE}{Mat.GROUP_ROW}:{MATRIX_LAST_TITLE}{HEADER_ROW}"
    clear_when_blank(ws, header_cells, f"{first}${HEADER_ROW}")
    ws.conditional_formatting.add(f"{Mat.MISSING}{FIRST_ROW}:{Mat.MISSING}{PEOPLE_LAST}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(f"{Mat.ATTENTION}{FIRST_ROW}:{Mat.ATTENTION}{PEOPLE_LAST}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color=ORANGE_TEXT, bold=True)))


def create(wb):
    ws = wb.create_sheet(Sheet.MATRIX)
    title_block(ws, "Matrix", "Everyone vs. every title at a glance. Read-only - driven by Inventory and Role "
                              "Defaults. Use the filter arrows to narrow by department or role.")
    ws.sheet_properties.tabColor = TAB_REPORT
    write_headers(ws, LEFT_COLS, TITLE_HEADER_HEIGHT)
    _write_title_headers(ws)
    _write_body(ws)
    _add_formatting(ws)
    ws.freeze_panes = f"{MATRIX_FIRST_TITLE}{FIRST_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{MATRIX_LAST_TITLE}{PEOPLE_LAST}"
    page_setup(ws, f"{Mat.GROUP_ROW}:{HEADER_ROW}")
    protect_report(ws)
