"""Look and feel shared by every sheet. The whole theme (fonts, colours, sizes) is the block at the top.

Black type, light-rose headers, brass rules, hairline row dividers. Change the constants to re-theme the workbook.
"""
from dataclasses import dataclass

from openpyxl.formatting.rule import CellIsRule, FormulaRule, Rule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat

from layout import FIRST_ROW, HEADER_ROW, STALE_DAYS, WARN_DAYS, col_index, col_letter
from vocab import STATUS_STYLE, Applicability

# ---- theme -------------------------------------------------------------------------------------------------------
FONT_NAME = "Times New Roman"
SIZE_BODY = 12
SIZE_SMALL = 11          # notes, group labels and the Matrix grid
SIZE_HEADING = 14
SIZE_TITLE = 20
SIZE_FIGURE = 18         # the big numbers on Dashboard and Person View

INK = "000000"
ROSE = "F2C9C4"          # header of a column you type in
ROSE_PALE = "FBEDEB"
PARCHMENT = "F5F0E6"     # body of a calculated column
PARCHMENT_DARK = "E6DDCB"  # header of a calculated column
BRASS = "B08A3E"
BRASS_DARK = "7A5C1E"
RULE = "CDBFA6"          # hairline between rows
MUTED = "6B6254"
CRIMSON = "A61C1C"       # date already past, low coverage
RUST = "8A4B08"          # date inside the warning window
BAR = "D4B572"

TAB_INPUT, TAB_REFERENCE, TAB_REPORT, TAB_GUIDE = BRASS, "262626", "D98880", "8C8C8C"

WIDTH_SCALE = 1.2        # Times New Roman 12pt is wider than the Normal-style font the width unit is measured in
HEADER_HEIGHT = 48
SYMBOL_FORMAT_ID_BASE = 200   # custom number-format ids must be unique and >= 164

HAIRLINE = Side(style="thin", color=RULE)
BORDER = Border(bottom=HAIRLINE)
GRID = Border(left=HAIRLINE, right=HAIRLINE, top=HAIRLINE, bottom=HAIRLINE)
GROUP_EDGE = Side(style="medium", color=BRASS)
NO_BORDER = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))
HEADER_BORDER = Border(top=Side(style="thin", color=INK), bottom=Side(style="medium", color=BRASS))

APPLICABILITY_COLOURS = {   # applicability -> (fill, font)
    Applicability.REQUIRED: (ROSE, "7A1010"),
    Applicability.OPTIONAL: (PARCHMENT_DARK, "5C4A1E"),
    Applicability.NOT_APPLICABLE: ("DAD6CF", "6E6A63"),
}


@dataclass(frozen=True)
class Col:
    """One table column: where it is, what it says, how wide, and how its cells look."""
    letter: str
    title: str
    width: float
    calc: bool = False
    fmt: str | None = None
    align: str = "left"
    bold: bool = False


def fill(hex_colour):
    return PatternFill("solid", start_color=hex_colour, end_color=hex_colour)


def font(size=SIZE_BODY, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic, underline=underline)


def scaled(width):
    """A column width in the units the sheets use, adjusted for the theme font."""
    return width * WIDTH_SCALE


def title_block(ws, title, subtitle):
    """Sheet title in A1 and a one-line description in A2; gridlines off."""
    ws["A1"] = title.upper()
    ws["A1"].font = font(SIZE_HEADING, True)
    ws["A2"] = subtitle
    ws["A2"].font = font(SIZE_SMALL, color=MUTED, italic=True)
    ws.row_dimensions[1].height = 26
    ws.sheet_view.showGridLines = False


def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = font(SIZE_SMALL, color=MUTED)


def header_cell(ws, row, column, text, calc=False, width=None):
    cell = ws.cell(row=row, column=column, value=text)
    cell.font = font(bold=True, italic=calc)
    cell.fill = fill(PARCHMENT_DARK if calc else ROSE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = HEADER_BORDER
    if width:
        ws.column_dimensions[col_letter(column)].width = scaled(width)
    return cell


def write_headers(ws, cols, height=HEADER_HEIGHT):
    """Header row for a list of `Col`."""
    for col in cols:
        header_cell(ws, HEADER_ROW, col_index(col.letter), col.title, col.calc, col.width)
    ws.row_dimensions[HEADER_ROW].height = height


def style_cells(ws, cells, calc=False, fmt=None, align="left", bold=False, size=SIZE_BODY):
    """Body style for a range such as 'A6:M40'. Cells are locked until `unlock` is called on them."""
    for row in ws[cells]:
        for cell in row:
            cell.font = font(size, bold)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if calc:
                cell.fill = fill(PARCHMENT)
            if fmt:
                cell.number_format = fmt


def unlock(ws, cells):
    """Keep a range editable once the sheet is protected."""
    for row in ws[cells]:
        for cell in row:
            cell.protection = Protection(locked=False)


def style_columns(ws, cols, last_row):
    """Apply each column's body style from FIRST_ROW to `last_row`; typed-in columns are unlocked."""
    for col in cols:
        cells = f"{col.letter}{FIRST_ROW}:{col.letter}{last_row}"
        style_cells(ws, cells, col.calc, col.fmt, col.align, col.bold)
        if not col.calc:
            unlock(ws, cells)


def status_colours(ws, cells):
    """Colour cells holding a status name and put its symbol in front (display only; the value is unchanged)."""
    for index, (status, style) in enumerate(STATUS_STYLE.items()):
        symbol_format = NumberFormat(numFmtId=SYMBOL_FORMAT_ID_BASE + index, formatCode=f'"{style.symbol} "@')
        dxf = DifferentialStyle(font=Font(color=style.font, bold=style.bold), fill=fill(style.fill),
                                numFmt=symbol_format)
        ws.conditional_formatting.add(cells, Rule(type="cellIs", operator="equal", formula=[f'"{status}"'], dxf=dxf))


def applicability_colours(ws, cells):
    for name, (background, text) in APPLICABILITY_COLOURS.items():
        ws.conditional_formatting.add(cells, CellIsRule(
            operator="equal", formula=[f'"{name}"'], fill=fill(background), font=Font(color=text)))


def text_flags(ws, cells, flags):
    """Colour cells whose text equals a key of `flags` (text -> font colour); all bold."""
    for text, colour in flags.items():
        ws.conditional_formatting.add(cells, CellIsRule(
            operator="equal", formula=[f'"{text}"'], font=Font(color=colour, bold=True)))


def _anchor(cells):
    """Column-absolute reference to the first cell of a range: 'O6:O99' -> '$O6'."""
    first = cells.split(":")[0].replace("$", "")
    letters = first.rstrip("0123456789")
    return f"${letters}{first[len(letters):]}"


def date_flags(ws, cells):
    """Crimson when the date has passed, rust when it falls inside the warning window."""
    ref = _anchor(cells)
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}<TODAY())'], font=Font(color=CRIMSON, bold=True)))
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}<=TODAY()+{WARN_DAYS})'], font=Font(color=RUST, bold=True)))


def stale_flag(ws, cells):
    """Crimson when the 'days since verified' number exceeds the stale threshold."""
    ref = _anchor(cells)
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}>{STALE_DAYS})'], font=Font(color=CRIMSON, bold=True)))


def clear_when_blank(ws, cells, key_cell):
    """Remove fill and borders where `key_cell` is blank, so unused capacity stays invisible."""
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'{key_cell}=""'], fill=fill("FFFFFF"), border=NO_BORDER))


def page_setup(ws, repeat_rows=None, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    if repeat_rows:
        ws.print_title_rows = repeat_rows


def protect_sheet(ws):
    """Lock formulas and structure (no password). Unlocked cells stay editable; filtering and resizing still work.

    Inserting, deleting and sorting rows are blocked because other sheets find data by row position.
    """
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.autoFilter = False
