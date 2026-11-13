"""Dashboard sheet: a strip of headline figures plus the three roll-up tables. Read-only; computed from the Matrix."""
from openpyxl.styles import Alignment, Border, Side

import refs
from dashboard_tables import PERCENT_FORMAT, department_table, group_table, software_table, write_table
from layout import N_SUMMARY_ROWS, Cat, Ppl, Sheet, col_index
from styles import (
    BRASS,
    BRASS_DARK,
    CRIMSON,
    INK,
    SIZE_FIGURE,
    SIZE_SMALL,
    TAB_REPORT,
    font,
    note,
    page_setup,
    protect_sheet,
    title_block,
)
from vocab import ATTENTION, HAVE, Active, Status

LABEL_ROW, VALUE_ROW, NOTE_ROW = 4, 5, 6
VALUE_HEIGHT = 34
FIRST_TABLE_HEADER = 9
TABLE_GAP = 3             # blank row + title row + header row between tables
FROZEN_ROWS = "A4"

# label, (first col, last col), colour of the figure
FIGURES = [
    ("PEOPLE", ("A", "A"), INK),
    ("SOFTWARE TITLES", ("B", "C"), INK),
    ("INSTALLED (HAVE)", ("D", "E"), INK),
    ("MISSING", ("F", "G"), CRIMSON),
    ("NEEDS ATTENTION", ("H", "I"), INK),
    ("COVERAGE", ("J", "K"), INK),
]
EXPLANATION = ('Coverage = Have ÷ (Have + Missing + Pending). "Have" counts Installed, Outdated, Expiring, '
               'License Issue and Verify. "Needs Attention" counts Outdated, Expiring, License Issue, Verify, '
               'Missing, Review and Unknown. Retired people are not counted.')


def _count(statuses):
    body = refs.matrix_body()
    return "+".join(f'COUNTIF({body},"{status}")' for status in statuses)


def _figure_formulas():
    """Formula (with '=') for each figure, in FIGURES order."""
    body = refs.matrix_body()
    have = _count(HAVE)
    return [
        f'=COUNTIFS({refs.ppl(Ppl.NAME)},"?*",{refs.ppl(Ppl.ACTIVE)},"<>{Active.NO}")',
        f'=COUNTIF({refs.cat(Cat.NAME)},"?*")',
        f"={have}",
        f'=COUNTIF({body},"{Status.MISSING}")',
        f"={_count(ATTENTION)}",
        f'=IFERROR(({have})/(({have})+COUNTIF({body},"{Status.MISSING}")+COUNTIF({body},"{Status.PENDING}")),"—")',
    ]


def _write_figures(ws):
    rule = Border(top=Side(style="medium", color=BRASS))
    for (label, (first, last), colour), formula in zip(FIGURES, _figure_formulas(), strict=True):
        if first != last:
            ws.merge_cells(f"{first}{LABEL_ROW}:{last}{LABEL_ROW}")
            ws.merge_cells(f"{first}{VALUE_ROW}:{last}{VALUE_ROW}")
        ws[f"{first}{LABEL_ROW}"] = label
        ws[f"{first}{VALUE_ROW}"] = formula
        for column in range(col_index(first), col_index(last) + 1):
            top, value = ws.cell(row=LABEL_ROW, column=column), ws.cell(row=VALUE_ROW, column=column)
            top.font, value.font = font(SIZE_SMALL, True, BRASS_DARK), font(SIZE_FIGURE, True, colour)
            top.border = rule
            top.alignment = value.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws[f"{FIGURES[-1][1][0]}{VALUE_ROW}"].number_format = PERCENT_FORMAT
    ws.row_dimensions[VALUE_ROW].height = VALUE_HEIGHT
    note(ws, f"A{NOTE_ROW}", EXPLANATION)


def _tables():
    """The three tables with their header rows stacked top to bottom."""
    department_header = FIRST_TABLE_HEADER + N_SUMMARY_ROWS + TABLE_GAP
    software_header = department_header + N_SUMMARY_ROWS + TABLE_GAP
    software = software_table(software_header)
    return [group_table(FIRST_TABLE_HEADER, software), department_table(department_header), software]


def create(wb):
    ws = wb.create_sheet(Sheet.DASHBOARD)
    title_block(ws, "Dashboard", "Roll-up of the whole population. Read-only - calculated from the Matrix.")
    ws.sheet_properties.tabColor = TAB_REPORT
    _write_figures(ws)
    for table in _tables():
        write_table(ws, table)
    ws.freeze_panes = FROZEN_ROWS
    page_setup(ws)
    protect_sheet(ws)
