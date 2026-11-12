"""Start Here sheet: purpose, five-step workflow, clickable sheet index and a few tips."""
from openpyxl.styles import Alignment

from layout import N_PEOPLE, N_RECORDS, N_ROLES, N_TITLES, Sheet
from styles import NAVY, SLATE, TAB_GUIDE, fill, font, page_setup

TEMPLATE_VERSION = "1.0"
MARGIN_WIDTH, NAME_WIDTH, TEXT_WIDTH = 3, 24, 90
LINK_COLOUR = "1F4E9C"
FIRST_SECTION_ROW = 6

STEPS = [
    ("1. Maintain Catalog", "List each software title once: group/suite, approved version, license model, server, "
                            "seats and expiry."),
    ("2. Add People", "One row each, with a department and role."),
    ("3. Set Role Defaults", "Mark each role's titles Required / Optional / Not Applicable."),
    ("4. Record installs", "On Inventory, add one row per person + software that is installed, pending or needs a "
                           "note. No row means not installed."),
    ("5. Read the reports", "Matrix, Person View and Dashboard update automatically."),
]
SHEET_GUIDE = [
    (Sheet.LEGEND, "Every status and colour explained."),
    (Sheet.INVENTORY, "DATA ENTRY - one row per person + software."),
    (Sheet.MATRIX, "Everyone x every title at a glance (read-only)."),
    (Sheet.PERSON_VIEW, "Pick a person to see their full list (read-only)."),
    (Sheet.DASHBOARD, "Coverage by group, department and software (read-only)."),
    (Sheet.CATALOG, "Software titles, licenses, seats, servers, expiry."),
    (Sheet.PEOPLE, "Roster with department and role."),
    (Sheet.ROLES, "Required / Optional / N/A by role."),
    (Sheet.LISTS, "Dropdown lists and warning thresholds."),
]
TIPS = [
    f"Capacity: {N_PEOPLE} people, {N_TITLES} software titles, {N_ROLES} roles, {N_RECORDS:,} inventory rows. "
    "Formulas and dropdowns are pre-filled to that size.",
    "Navy headers = you type, grey = calculated. Don't overwrite grey cells.",
    "Matrix, Person View and Dashboard are protected (no password) so formulas can't be broken; "
    "use Review > Unprotect Sheet if you need to.",
    "Rows not verified within the 'stale' window (Lists sheet) show as Verify.",
    "Works in desktop Excel 2013 or later; no macros.",
]


def _section(ws, row, text):
    ws[f"B{row}"] = text
    ws[f"B{row}"].font = font(13, True, NAVY)
    return row + 1


def _steps(ws, row):
    for step, text in STEPS:
        ws[f"B{row}"], ws[f"C{row}"] = step, text
        ws[f"B{row}"].font = font(10, True)
        ws[f"B{row}"].alignment = Alignment(vertical="top")
        ws[f"C{row}"].font = font()
        ws[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    return row + 1


def _sheet_index(ws, row):
    for name, description in SHEET_GUIDE:
        ws[f"B{row}"], ws[f"C{row}"] = name, description
        ws[f"B{row}"].hyperlink = f"#'{name}'!A1"
        ws[f"B{row}"].font = font(10, True, LINK_COLOUR, underline="single")
        ws[f"C{row}"].font = font()
        row += 1
    return row + 1


def _tips(ws, row):
    for tip in TIPS:
        ws[f"B{row}"] = "•"
        ws[f"B{row}"].alignment = Alignment(horizontal="right", vertical="top")
        ws[f"C{row}"] = tip
        ws[f"C{row}"].font = font()
        ws[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    return row


def fill_start(ws):
    ws.title = Sheet.START
    ws.sheet_properties.tabColor = TAB_GUIDE
    ws.sheet_view.showGridLines = False
    for letter, width in (("A", MARGIN_WIDTH), ("B", NAME_WIDTH), ("C", TEXT_WIDTH)):
        ws.column_dimensions[letter].width = width
    ws["B2"] = "Software Inventory"
    ws["B2"].font = font(26, True, NAVY)
    ws["B3"] = "Who has what software installed, what they're missing, and what doesn't apply to them."
    ws["B3"].font = font(12, False, SLATE, italic=True)
    ws.row_dimensions[2].height = 40
    for letter in "BC":
        ws[f"{letter}4"].fill = fill(NAVY)
    ws.row_dimensions[4].height = 4
    row = _steps(ws, _section(ws, FIRST_SECTION_ROW, "How it works"))
    row = _sheet_index(ws, _section(ws, row, "Sheets (click to jump)"))
    row = _tips(ws, _section(ws, row, "Good to know"))
    ws[f"B{row + 1}"] = f"Template version {TEMPLATE_VERSION}"
    ws[f"B{row + 1}"].font = font(9, False, SLATE, italic=True)
    page_setup(ws, landscape=False)
