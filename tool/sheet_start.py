"""About sheet: what the workbook is for, how to use it, the rules that keep it correct, and an index of the sheets."""
from openpyxl.styles import Alignment, Border, Side

from layout import N_PEOPLE, N_RECORDS, N_ROLES, N_TITLES, Sheet
from styles import BRASS, BRASS_DARK, INK, MUTED, SIZE_BODY, SIZE_HEADING, TAB_GUIDE, font, page_setup, scaled

MARGIN_WIDTH, NAME_WIDTH, TEXT_WIDTH = 3, 24, 95
LINK_COLOUR = "7A5C1E"
TITLE_SIZE = 24
FIRST_SECTION_ROW = 6

WORKFLOW = [
    ("Catalog", "List each software title once: group, approved version, license model, server, seats, expiry."),
    ("People", "One row per person with department and role. Set Active to No when someone leaves."),
    ("Role Defaults", "For each role, mark every title Required, Optional or Not Applicable."),
    ("Inventory", "Already laid out: every person has a row for every title. Fill in only what is installed."),
    ("Reports", "Matrix, Person View and Dashboard update on their own."),
]
RULES = [
    "Add people and titles in the next empty row of People or Catalog. Never insert, delete or sort rows there; "
    "other sheets find data by row position.",
    "Retire a person (Active = No) or leave a title's row in place instead of deleting it.",
    "Leave Install State blank when nothing is installed - Missing, Not Installed or N/A follows from the role.",
    "Use Applicability Override on Inventory for one-off exceptions; change Role Defaults for everyone in a role.",
    "People, Catalog, Role Defaults and the report sheets are protected without a password so the structure "
    "survives; Review > Unprotect Sheet if you need to.",
    f"Room for {N_PEOPLE} people, {N_TITLES} titles, {N_ROLES} roles and {N_RECORDS:,} Inventory rows.",
]
SHEET_GUIDE = [
    (Sheet.LEGEND, "What every status and colour means."),
    (Sheet.INVENTORY, "Data entry - install details for each person and title."),
    (Sheet.MATRIX, "Everyone against every title (read-only)."),
    (Sheet.PERSON_VIEW, "One person's full list (read-only)."),
    (Sheet.DASHBOARD, "Coverage by group, department and software (read-only)."),
    (Sheet.CATALOG, "Titles, licenses, seats, servers, expiry."),
    (Sheet.PEOPLE, "Roster with department, role and active flag."),
    (Sheet.ROLES, "Required / Optional / Not Applicable by role."),
    (Sheet.LISTS, "Dropdown lists and warning thresholds."),
]


def _section(ws, row, text):
    ws[f"B{row}"] = text.upper()
    ws[f"B{row}"].font = font(SIZE_BODY, True, BRASS_DARK)
    for letter in "BC":
        ws[f"{letter}{row}"].border = Border(bottom=Side(style="thin", color=BRASS))
    return row + 1


def _two_columns(ws, row, pairs, link=False):
    for name, text in pairs:
        ws[f"B{row}"], ws[f"C{row}"] = name, text
        colour = LINK_COLOUR if link else INK
        ws[f"B{row}"].font = font(SIZE_BODY, True, colour, underline="single" if link else None)
        ws[f"B{row}"].alignment = Alignment(vertical="top")
        ws[f"C{row}"].font = font()
        ws[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        if link:
            ws[f"B{row}"].hyperlink = f"#'{name}'!A1"
        row += 1
    return row + 1


def _bullets(ws, row, lines):
    for line in lines:
        ws[f"B{row}"] = "–"
        ws[f"B{row}"].alignment = Alignment(horizontal="right", vertical="top")
        ws[f"C{row}"] = line
        ws[f"C{row}"].font = font()
        ws[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    return row + 1


def fill_start(ws):
    ws.title = Sheet.START
    ws.sheet_properties.tabColor = TAB_GUIDE
    ws.sheet_view.showGridLines = False
    for letter, width in (("A", MARGIN_WIDTH), ("B", NAME_WIDTH), ("C", TEXT_WIDTH)):
        ws.column_dimensions[letter].width = scaled(width)
    ws["B2"] = "SOFTWARE INVENTORY"
    ws["B2"].font = font(TITLE_SIZE, True)
    ws["B3"] = "Who has what installed, what is missing, and what does not apply."
    ws["B3"].font = font(SIZE_HEADING - 2, False, MUTED, italic=True)
    ws.row_dimensions[2].height = 36
    row = _two_columns(ws, _section(ws, FIRST_SECTION_ROW, "Workflow"), WORKFLOW)
    row = _bullets(ws, _section(ws, row, "Rules"), RULES)
    _two_columns(ws, _section(ws, row, "Sheets"), SHEET_GUIDE, link=True)
    page_setup(ws, landscape=False)
