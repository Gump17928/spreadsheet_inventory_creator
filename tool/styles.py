"""Look and feel shared by every sheet: palette, fonts, borders, header rows, conditional-format helpers."""
from dataclasses import dataclass

from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from layout import FIRST_ROW, HEADER_ROW, STALE_DAYS, WARN_DAYS, col_index, col_letter
from vocab import STATUS_STYLE, Applicability

NAVY = "1F3A5F"       # headers you type under, titles
SLATE = "5B6B7F"      # headers of calculated columns, subtitles
PALE = "F3F6FA"
INK = "263238"
CALC_FILL = "EEF1F5"
BAND_FILL = "F8FAFD"
FONT_NAME = "Calibri"
RED_TEXT = "C00000"       # date already past
ORANGE_TEXT = "9C4A00"    # date inside the warning window

TAB_INPUT, TAB_REFERENCE, TAB_REPORT, TAB_GUIDE = "70AD47", "4F81BD", "8E7CC3", NAVY

_EDGE = Side(style="thin", color="D0D7E2")
BORDER = Border(left=_EDGE, right=_EDGE, top=_EDGE, bottom=_EDGE)
NO_BORDER = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

APPLICABILITY_COLOURS = {
    Applicability.REQUIRED: ("FFC7CE", "9C0006"),
    Applicability.OPTIONAL: ("DDEBF7", "1F4E78"),
    Applicability.NOT_APPLICABLE: ("D9D9D9", "7F7F7F"),
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


def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic, underline=underline)


def title_block(ws, title, subtitle):
    """Sheet title in A1 and a one-line description in A2; gridlines off."""
    ws["A1"] = title
    ws["A1"].font = font(18, True, NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = font(10, False, SLATE, italic=True)
    ws.row_dimensions[1].height = 28
    ws.sheet_view.showGridLines = False


def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = font(9, False, SLATE)


def header_cell(ws, row, column, text, calc=False, width=None):
    cell = ws.cell(row=row, column=column, value=text)
    cell.font = font(10, True, "FFFFFF")
    cell.fill = fill(SLATE if calc else NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
    if width:
        ws.column_dimensions[col_letter(column)].width = width
    return cell


def write_headers(ws, cols, height=42):
    """Header row for a list of `Col`."""
    for col in cols:
        header_cell(ws, HEADER_ROW, col_index(col.letter), col.title, col.calc, col.width)
    ws.row_dimensions[HEADER_ROW].height = height


def style_cells(ws, cells, calc=False, fmt=None, align="left", bold=False):
    """Plain bordered body style for a range such as 'A6:M40'."""
    for row in ws[cells]:
        for cell in row:
            cell.font = font(bold=bold)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if calc:
                cell.fill = fill(CALC_FILL)
            if fmt:
                cell.number_format = fmt


def style_columns(ws, cols, last_row):
    """Apply each column's body style from FIRST_ROW to `last_row`."""
    for col in cols:
        style_cells(ws, f"{col.letter}{FIRST_ROW}:{col.letter}{last_row}", col.calc, col.fmt, col.align, col.bold)


def band(ws, cells):
    """Light alternating-row shading."""
    ws.conditional_formatting.add(cells, FormulaRule(formula=["MOD(ROW(),2)=0"], fill=fill(BAND_FILL)))


def status_colours(ws, cells):
    """Colour cells holding a status name."""
    for status, style in STATUS_STYLE.items():
        ws.conditional_formatting.add(cells, CellIsRule(
            operator="equal", formula=[f'"{status}"'], fill=fill(style.fill),
            font=Font(color=style.font, bold=style.bold)))


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
    """Red when the date has passed, orange when it falls inside the warning window."""
    ref = _anchor(cells)
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}<TODAY())'], font=Font(color=RED_TEXT, bold=True)))
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}<=TODAY()+{WARN_DAYS})'], font=Font(color=ORANGE_TEXT, bold=True)))


def stale_flag(ws, cells):
    """Red when the 'days since verified' number exceeds the stale threshold."""
    ref = _anchor(cells)
    ws.conditional_formatting.add(cells, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}>{STALE_DAYS})'], font=Font(color=RED_TEXT, bold=True)))


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


def protect_report(ws):
    """Lock a report sheet (no password) while still letting users filter and resize."""
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.autoFilter = False
    ws.protection.sort = False
