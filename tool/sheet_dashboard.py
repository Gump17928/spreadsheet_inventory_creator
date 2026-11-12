"""Dashboard sheet: headline tiles plus the three roll-up tables. Read-only; computed from the Matrix."""
from openpyxl.styles import Alignment

import refs
from dashboard_tables import PERCENT_FORMAT, department_table, group_table, software_table, write_table
from layout import N_SUMMARY_ROWS, Cat, Ppl, Sheet, col_index
from styles import TAB_REPORT, fill, font, note, page_setup, protect_report, title_block
from vocab import ATTENTION, HAVE, Status

TILE_LABEL_ROW, TILE_VALUE_ROW, TILE_NOTE_ROW = 4, 5, 6
TILE_VALUE_HEIGHT = 38
FIRST_TABLE_HEADER = 9
TABLE_GAP = 3             # blank row + title row + header row between tables
FROZEN_ROWS = "A4"

# label, (first col, last col), fill, text colour
TILE_LAYOUT = [
    ("PEOPLE", ("A", "A"), "DDEBF7", "1F4E78"),
    ("SOFTWARE TITLES", ("B", "C"), "DDEBF7", "1F4E78"),
    ("INSTALLED (HAVE)", ("D", "E"), "D9EAD3", "1E5631"),
    ("MISSING", ("F", "G"), "FFC7CE", "9C0006"),
    ("NEEDS ATTENTION", ("H", "I"), "FFE699", "7F6000"),
    ("COVERAGE", ("J", "K"), "E4DFEC", "5B3F8C"),
]
EXPLANATION = ('Coverage = Have ÷ (Have + Missing + Pending). "Have" counts Installed, Outdated, Expiring, '
               'License Issue and Verify. "Needs Attention" counts Outdated, Expiring, License Issue, Verify, '
               'Missing, Review and Unknown.')


def _count(statuses):
    body = refs.matrix_body()
    return "+".join(f'COUNTIF({body},"{status}")' for status in statuses)


def _tile_formulas():
    """Formula (with '=') for each tile, in TILE_LAYOUT order."""
    body = refs.matrix_body()
    have = _count(HAVE)
    return [
        f'=COUNTIF({refs.ppl(Ppl.NAME)},"?*")',
        f'=COUNTIF({refs.cat(Cat.NAME)},"?*")',
        f"={have}",
        f'=COUNTIF({body},"{Status.MISSING}")',
        f"={_count(ATTENTION)}",
        f'=IFERROR(({have})/(({have})+COUNTIF({body},"{Status.MISSING}")+COUNTIF({body},"{Status.PENDING}")),"—")',
    ]


def _write_tiles(ws):
    for (label, (first, last), background, text), formula in zip(TILE_LAYOUT, _tile_formulas(), strict=True):
        if first != last:
            ws.merge_cells(f"{first}{TILE_LABEL_ROW}:{last}{TILE_LABEL_ROW}")
            ws.merge_cells(f"{first}{TILE_VALUE_ROW}:{last}{TILE_VALUE_ROW}")
        ws[f"{first}{TILE_LABEL_ROW}"] = label
        ws[f"{first}{TILE_VALUE_ROW}"] = formula
        for column in range(col_index(first), col_index(last) + 1):
            for row, size in ((TILE_LABEL_ROW, 9), (TILE_VALUE_ROW, 22)):
                cell = ws.cell(row=row, column=column)
                cell.fill = fill(background)
                cell.font = font(size, True, text)
                cell.alignment = Alignment(horizontal="center", vertical="center")
    ws[f"{TILE_LAYOUT[-1][1][0]}{TILE_VALUE_ROW}"].number_format = PERCENT_FORMAT
    ws.row_dimensions[TILE_VALUE_ROW].height = TILE_VALUE_HEIGHT
    note(ws, f"A{TILE_NOTE_ROW}", EXPLANATION)


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
    _write_tiles(ws)
    for table in _tables():
        write_table(ws, table)
    ws.freeze_panes = FROZEN_ROWS
    page_setup(ws)
    protect_report(ws)
