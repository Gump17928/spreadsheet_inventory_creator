"""Matrix sheet: every person against every title, colour-coded. Read-only; all cells are formulas."""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font

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
    inventory_row,
)
from lookups import person_at
from refs import at
from styles import (
    BRASS_DARK,
    CRIMSON,
    GRID,
    GROUP_EDGE,
    HAIRLINE,
    PARCHMENT_DARK,
    RUST,
    SIZE_SMALL,
    TAB_REPORT,
    Col,
    clear_when_blank,
    fill,
    font,
    header_cell,
    page_setup,
    protect_sheet,
    status_colours,
    style_cells,
    title_block,
    write_headers,
)
from vocab import ATTENTION, HAVE, Active, Status

GROUP_ROW_HEIGHT = 36
TITLE_HEADER_HEIGHT = 66
TITLE_COL_WIDTH = 15
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
    """Person, department, role, computer and the three counters for one Matrix row (= one People row)."""
    person = f"${Mat.PERSON}{row}"
    row_range = f"${MATRIX_FIRST_TITLE}{row}:${MATRIX_LAST_TITLE}{row}"
    retired = f'{person_at(Ppl.ACTIVE, row)}="{Active.NO}"'
    name = person_at(Ppl.NAME, row)
    return {
        Mat.PERSON: f'=IF(OR({name}="",{retired}),"",{name})',
        Mat.DEPARTMENT: f'=IF({person}="","",{person_at(Ppl.DEPARTMENT, row)}&"")',
        Mat.ROLE: f'=IF({person}="","",{person_at(Ppl.ROLE, row)}&"")',
        Mat.COMPUTER: f'=IF({person}="","",{person_at(Ppl.COMPUTER, row)}&"")',
        Mat.HAVE: f'=IF({person}="","",{_count(HAVE, row_range)})',
        Mat.MISSING: f'=IF({person}="","",COUNTIF({row_range},"{Status.MISSING}"))',
        Mat.ATTENTION: f'=IF({person}="","",{_count(ATTENTION, row_range)})',
    }


def cell_formula(row, slot):
    """Status of one person for one title, read from that person's Inventory row for the title."""
    title = f"{Mat.title_col(slot)}${HEADER_ROW}"
    status = at(Sheet.INVENTORY, Inv.STATUS, inventory_row(row - FIRST_ROW, slot))
    return f'=IF(OR(${Mat.PERSON}{row}="",{title}=""),"",{status})'


def _catalog_mirror(column, slot):
    """Formula that shows a Catalog value for slot `slot`, or blank when the slot is empty."""
    name = at(Sheet.CATALOG, Cat.NAME, FIRST_ROW + slot)
    return f'=IF({name}="","",{at(Sheet.CATALOG, column, FIRST_ROW + slot)})'


def _write_title_headers(ws):
    """Group names above, title names below; both mirror the Catalog order."""
    for slot in range(N_TITLES):
        index = col_index(Mat.title_col(slot))
        group = ws.cell(row=Mat.GROUP_ROW, column=index, value=_catalog_mirror(Cat.GROUP, slot))
        group.font = font(SIZE_SMALL, True, BRASS_DARK)
        group.fill = fill(PARCHMENT_DARK)
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
    style_cells(ws, BODY, align="center", size=SIZE_SMALL)


def _add_formatting(ws):
    """Status colours; gridlines and brass group dividers appear only where there is data."""
    status_colours(ws, BODY)
    first = MATRIX_FIRST_TITLE
    used = f'AND(${Mat.PERSON}{FIRST_ROW}<>"",{first}${HEADER_ROW}<>"")'
    ws.conditional_formatting.add(BODY, FormulaRule(
        formula=[f'AND({used},{first}${Mat.GROUP_ROW}<>{Mat.ATTENTION}${Mat.GROUP_ROW})'],
        border=Border(left=GROUP_EDGE, right=HAIRLINE, top=HAIRLINE, bottom=HAIRLINE)))
    ws.conditional_formatting.add(BODY, FormulaRule(formula=[used], border=GRID))
    clear_when_blank(ws, f"{Mat.PERSON}{FIRST_ROW}:{MATRIX_LAST_TITLE}{PEOPLE_LAST}", f"${Mat.PERSON}{FIRST_ROW}")
    header_cells = f"{MATRIX_FIRST_TITLE}{Mat.GROUP_ROW}:{MATRIX_LAST_TITLE}{PEOPLE_LAST}"
    clear_when_blank(ws, header_cells, f"{first}${HEADER_ROW}")
    ws.conditional_formatting.add(f"{Mat.MISSING}{FIRST_ROW}:{Mat.MISSING}{PEOPLE_LAST}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color=CRIMSON, bold=True)))
    ws.conditional_formatting.add(f"{Mat.ATTENTION}{FIRST_ROW}:{Mat.ATTENTION}{PEOPLE_LAST}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color=RUST, bold=True)))


def create(wb):
    ws = wb.create_sheet(Sheet.MATRIX)
    title_block(ws, "Matrix", "Everyone against every title. Read-only - calculated from Inventory. "
                              "Retired people are left out. Use the filter arrows to narrow by department or role.")
    ws.sheet_properties.tabColor = TAB_REPORT
    write_headers(ws, LEFT_COLS, TITLE_HEADER_HEIGHT)
    _write_title_headers(ws)
    _write_body(ws)
    _add_formatting(ws)
    ws.freeze_panes = f"{MATRIX_FIRST_TITLE}{FIRST_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{MATRIX_LAST_TITLE}{PEOPLE_LAST}"
    page_setup(ws, f"{Mat.GROUP_ROW}:{HEADER_ROW}")
    protect_sheet(ws)
