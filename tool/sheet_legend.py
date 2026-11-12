"""Legend sheet: what every status, state and colour means."""
from openpyxl.styles import Alignment

from layout import Sheet
from styles import BORDER, NAVY, PALE, TAB_GUIDE, fill, font, header_cell, page_setup, title_block
from vocab import STATUS_STYLE

STATUS_COL_WIDTH, MEANING_COL_WIDTH, ACTION_COL_WIDTH = 18, 70, 38
FIRST_ROW = 5

# (heading, [(term, meaning)])
REFERENCE_BLOCKS = [
    ("INSTALL STATE  (you choose, per Inventory row)", [
        ("Installed", "Software is on the person's computer."),
        ("Not Installed", "Row kept on purpose, e.g. to record an override or a note. Rows are optional here."),
        ("Pending", "Requested or approved but not installed yet."),
        ("Unknown", "Needs a physical or remote check.")]),
    ("APPLICABILITY  (from Role Defaults; override per person on Inventory)", [
        ("Required", "This role is expected to have it. Not installed -> Missing."),
        ("Optional", "Nice to have / used by some. Not installed -> Not Installed (no alarm)."),
        ("Not Applicable", "Not for this role. Not installed -> N/A. Installed -> Review.")]),
    ("LICENSE ASSIGNMENT  (per install)", [
        ("Named User Seat", "Subscription tied to the person."),
        ("Network (Floating)", "Checks out from a license server (host:port is on the Catalog)."),
        ("Standalone Key", "Perpetual or node-locked key on the machine."),
        ("Site / Enterprise", "Covered by a company-wide agreement."),
        ("Free / Not Required", "Open source or free tool."),
        ("Unlicensed", "No valid license -> License Issue."),
        ("Unknown", "Not checked yet -> Verify.")]),
    ("VERSION CHECK  (Installed Version vs. Catalog 'Approved Version')", [
        ("Current / Outdated", "Compared numerically when both are plain numbers (2026 vs 2025). "
                               "Caution: Excel reads 3.9 as higher than 3.12."),
        ("Check", "Text versions such as '2025 R2' must match the approved version exactly, or they are flagged."),
        ("n/a", "No approved version is set in the Catalog.")]),
    ("COLOURS", [
        ("Navy header", "You type or choose a value."),
        ("Grey header / cells", "Calculated automatically - don't type here."),
        ("Red date", "License date is in the past."),
        ("Orange date", "License expires within the warning window (Lists sheet)."),
        ("Green tab", "Data entry (Inventory)."),
        ("Blue tab", "Reference tables you maintain (Catalog, People, Role Defaults, Lists)."),
        ("Purple tab", "Reports, read-only (Matrix, Person View, Dashboard).")]),
]


def _status_table(ws):
    """The coloured status key; returns the next free row."""
    ws[f"A{FIRST_ROW - 2}"] = "STATUS  (calculated on Inventory, shown on Matrix / Person View / Dashboard)"
    ws[f"A{FIRST_ROW - 2}"].font = font(10, True, NAVY)
    for index, text in enumerate(("Status", "Meaning", "What to do"), start=1):
        header_cell(ws, FIRST_ROW - 1, index, text)
    row = FIRST_ROW
    for status, style in STATUS_STYLE.items():
        ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = status.value, style.meaning, style.action
        ws[f"A{row}"].fill = fill(style.fill)
        ws[f"A{row}"].font = font(10, True, style.font)
        ws[f"A{row}"].alignment = Alignment(horizontal="center")
        for letter in "ABC":
            ws[f"{letter}{row}"].border = BORDER
        for letter in "BC":
            ws[f"{letter}{row}"].font = font()
        row += 1
    return row + 1


def _reference_block(ws, row, heading, entries):
    """One heading plus its term/meaning rows; returns the next free row."""
    ws[f"A{row}"] = heading
    ws[f"A{row}"].font = font(10, True, NAVY)
    row += 1
    for term, meaning in entries:
        ws[f"A{row}"], ws[f"B{row}"] = term, meaning
        ws[f"A{row}"].fill = fill(PALE)
        ws[f"A{row}"].font = font(10, True)
        ws[f"B{row}"].font = font()
        for letter in "ABC":
            ws[f"{letter}{row}"].border = BORDER
        row += 1
    return row + 1


def create(wb):
    ws = wb.create_sheet(Sheet.LEGEND)
    title_block(ws, "Legend", "What every status, state and colour means.")
    ws.sheet_properties.tabColor = TAB_GUIDE
    ws.column_dimensions["A"].width = STATUS_COL_WIDTH
    ws.column_dimensions["B"].width = MEANING_COL_WIDTH
    ws.column_dimensions["C"].width = ACTION_COL_WIDTH
    row = _status_table(ws)
    for heading, entries in REFERENCE_BLOCKS:
        row = _reference_block(ws, row, heading, entries)
    page_setup(ws, landscape=False)
