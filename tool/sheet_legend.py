"""Key sheet: what every status, state and colour means."""
from openpyxl.styles import Alignment

from layout import Sheet
from styles import (
    BORDER,
    BRASS_DARK,
    PARCHMENT,
    SIZE_BODY,
    TAB_GUIDE,
    fill,
    font,
    header_cell,
    page_setup,
    scaled,
    title_block,
)
from vocab import STATUS_STYLE

STATUS_COL_WIDTH, MEANING_COL_WIDTH, ACTION_COL_WIDTH = 20, 78, 38
FIRST_ROW = 6

# (heading, [(term, meaning)])
REFERENCE_BLOCKS = [
    ("INSTALL STATE  (typed, per Inventory row)", [
        ("(blank)", "Not installed. The status then follows the person's applicability for the title."),
        ("Installed", "Software is on the person's computer."),
        ("Not Installed", "Same as blank; use it to say so explicitly."),
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
        ("Rose header", "You type or choose a value."),
        ("Tan italic header, shaded cells", "Calculated - don't type here."),
        ("Greyed-out row on Inventory", "Not Applicable to that person's role."),
        ("Crimson date", "License date is in the past."),
        ("Rust date", "License expires within the warning window (Lists sheet)."),
        ("Brass tab", "Data entry (Inventory)."),
        ("Black tab", "Reference tables you maintain (Catalog, People, Role Defaults, Lists)."),
        ("Rose tab", "Reports, read-only (Matrix, Person View, Dashboard).")]),
]


def _status_table(ws):
    """The coloured status key; returns the next free row."""
    ws[f"A{FIRST_ROW - 3}"] = "STATUS  (calculated on Inventory, shown on Matrix, Person View and Dashboard)"
    ws[f"A{FIRST_ROW - 3}"].font = font(SIZE_BODY, True, BRASS_DARK)
    for index, text in enumerate(("Status", "Meaning", "What to do"), start=1):
        header_cell(ws, FIRST_ROW - 1, index, text)
    row = FIRST_ROW
    for status, style in STATUS_STYLE.items():
        ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = f"{style.symbol} {status.value}", style.meaning, style.action
        ws[f"A{row}"].fill = fill(style.fill)
        ws[f"A{row}"].font = font(SIZE_BODY, True, style.font)
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
    ws[f"A{row}"].font = font(SIZE_BODY, True, BRASS_DARK)
    row += 1
    for term, meaning in entries:
        ws[f"A{row}"], ws[f"B{row}"] = term, meaning
        ws[f"A{row}"].fill = fill(PARCHMENT)
        ws[f"A{row}"].font = font(SIZE_BODY, True)
        ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws[f"B{row}"].font = font()
        for letter in "ABC":
            ws[f"{letter}{row}"].border = BORDER
        row += 1
    return row + 1


def create(wb):
    ws = wb.create_sheet(Sheet.LEGEND)
    title_block(ws, "Key", "What every status, state and colour means.")
    ws.sheet_properties.tabColor = TAB_GUIDE
    for letter, width in (("A", STATUS_COL_WIDTH), ("B", MEANING_COL_WIDTH), ("C", ACTION_COL_WIDTH)):
        ws.column_dimensions[letter].width = scaled(width)
    row = _status_table(ws)
    for heading, entries in REFERENCE_BLOCKS:
        row = _reference_block(ws, row, heading, entries)
    page_setup(ws, landscape=False)
