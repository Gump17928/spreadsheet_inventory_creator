"""Inventory sheet: one row for every person and every catalog title, generated from People and Catalog.

Only the install details are typed in; who and what each row is for is calculated, so nothing is ever created by hand.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

import inventory_formulas as formulas
import validation
from cell_values import number_if_plain
from layout import FIRST_ROW, HEADER_ROW, INV_LAST, Inv, Sheet, inventory_row
from styles import (
    MUTED,
    TAB_INPUT,
    Col,
    clear_when_blank,
    date_flags,
    fill,
    note,
    page_setup,
    protect_sheet,
    stale_flag,
    status_colours,
    style_columns,
    text_flags,
    title_block,
    write_headers,
)
from vocab import Applicability

DATE_FORMAT = "yyyy-mm-dd"
COLS = [
    Col(Inv.PERSON, "Person", 22, True), Col(Inv.COMPUTER, "Computer / Asset Tag", 17, True),
    Col(Inv.SOFTWARE, "Software", 24, True), Col(Inv.GROUP, "Group / Suite", 18, True),
    Col(Inv.STATE, "Install State", 14), Col(Inv.VERSION, "Installed Version", 12, align="center"),
    Col(Inv.INSTALLED_ON, "Install Date", 13, fmt=DATE_FORMAT, align="center"),
    Col(Inv.ASSIGNMENT, "License Assignment", 20),
    Col(Inv.EXPIRY_OVERRIDE, "License Expiry (override)", 14, fmt=DATE_FORMAT, align="center"),
    Col(Inv.APPL_OVERRIDE, "Applicability Override", 15),
    Col(Inv.VERIFIED_ON, "Last Verified", 13, fmt=DATE_FORMAT, align="center"), Col(Inv.VERIFIED_BY, "Verified By", 14),
    Col(Inv.NOTES, "Notes", 30),
    Col(Inv.APPLICABILITY, "Applicability (effective)", 15, True, align="center"),
    Col(Inv.EXPIRY, "License Expiry (effective)", 14, True, DATE_FORMAT, "center"),
    Col(Inv.VERSION_CHECK, "Version Check", 14, True, align="center"),
    Col(Inv.LICENSE_CHECK, "License Check", 14, True, align="center"),
    Col(Inv.DAYS_SINCE, "Days Since Verified", 13, True, "0", "center"),
    Col(Inv.STATUS, "STATUS", 16, True, align="center", bold=True),
    Col(Inv.LIVE, "Live", 6, True),
]
FORMULAS = {
    Inv.PERSON: formulas.person, Inv.COMPUTER: formulas.computer, Inv.SOFTWARE: formulas.software,
    Inv.GROUP: formulas.group, Inv.LIVE: formulas.live, Inv.APPLICABILITY: formulas.applicability,
    Inv.EXPIRY: formulas.expiry, Inv.VERSION_CHECK: formulas.version_check, Inv.LICENSE_CHECK: formulas.license_check,
    Inv.DAYS_SINCE: formulas.days_since_verified, Inv.STATUS: formulas.status,
}
VERSION_FLAGS = {"Outdated": "8A4B08", "Check": "8A4B08", "No version": "8A4B08"}
LICENSE_FLAGS = {"Expired": "A61C1C", "Unlicensed": "A61C1C", "Expiring": "8A4B08", "Unverified": "8A4B08"}
NOT_APPLICABLE_SHADE = "ECE8E1"
INPUT_RANGE = f"{Inv.STATE}{FIRST_ROW}:{Inv.NOTES}{INV_LAST}"
WHOLE_ROW = f"{Inv.PERSON}{FIRST_ROW}:{Inv.STATUS}{INV_LAST}"


def _cells(col):
    return f"{col}{FIRST_ROW}:{col}{INV_LAST}"


def _write_formulas(ws):
    for row in range(FIRST_ROW, INV_LAST + 1):
        for letter, build in FORMULAS.items():
            ws[f"{letter}{row}"] = build(row)


def _add_validation(ws):
    validation.list_rule(ws, _cells(Inv.STATE), "=InstallStates", "Install State",
                         "Leave blank if not installed. Otherwise Installed, Pending or Unknown.")
    validation.list_rule(ws, _cells(Inv.ASSIGNMENT), "=LicenseAssignList", "License",
                         "How this install is licensed. 'Unlicensed' triggers a License Issue.")
    validation.list_rule(ws, _cells(Inv.APPL_OVERRIDE), "=ApplicabilityList", "Override",
                         "Leave blank to use the role default. Set only for exceptions.")
    validation.date_rule(ws, _cells(Inv.INSTALLED_ON), "Install Date", "Date installed.", not_future=True)
    validation.date_rule(ws, _cells(Inv.EXPIRY_OVERRIDE), "License Expiry",
                         "Only fill in if this person's license differs from the Catalog expiry.")
    validation.date_rule(ws, _cells(Inv.VERIFIED_ON), "Last Verified",
                         "Date someone confirmed this on the machine.", not_future=True)
    validation.short_text_rule(ws, _cells(Inv.VERSION), "Version", "e.g. 2026, 2025 R2, 3.12.1")


def _add_formatting(ws):
    ws.conditional_formatting.add(INPUT_RANGE, FormulaRule(
        formula=[f'${Inv.APPLICABILITY}{FIRST_ROW}="{Applicability.NOT_APPLICABLE}"'],
        fill=fill(NOT_APPLICABLE_SHADE), font=Font(color=MUTED)))
    ws.conditional_formatting.add(WHOLE_ROW, FormulaRule(
        formula=[f'AND(${Inv.PERSON}{FIRST_ROW}<>"",${Inv.LIVE}{FIRST_ROW}="")'], font=Font(color=MUTED, italic=True)))
    status_colours(ws, _cells(Inv.STATUS))
    date_flags(ws, _cells(Inv.EXPIRY))
    stale_flag(ws, _cells(Inv.DAYS_SINCE))
    text_flags(ws, _cells(Inv.VERSION_CHECK), VERSION_FLAGS)
    text_flags(ws, _cells(Inv.LICENSE_CHECK), LICENSE_FLAGS)
    clear_when_blank(ws, WHOLE_ROW, f"${Inv.LIVE}{FIRST_ROW}")


def create(wb):
    ws = wb.create_sheet(Sheet.INVENTORY)
    title_block(ws, "Inventory", "One row per person and software title, filled in from People and Catalog. "
                                 "Type only in the rose-headed columns.")
    ws.sheet_properties.tabColor = TAB_INPUT
    note(ws, "A3", "Blank state = not installed: Required shows Missing, Optional shows Not Installed, Not Applicable "
                   "shows N/A. Filter Person to see one person; filter Software > Non-blanks to hide unused slots.")
    write_headers(ws, COLS)
    style_columns(ws, COLS, INV_LAST)
    _write_formulas(ws)
    _add_validation(ws)
    _add_formatting(ws)
    ws.column_dimensions[Inv.LIVE].hidden = True
    ws.freeze_panes = f"{Inv.GROUP}{FIRST_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{Inv.LAST_CALC}{INV_LAST}"
    page_setup(ws, f"{HEADER_ROW}:{HEADER_ROW}")
    protect_sheet(ws)


def write_records(ws, records, person_index, slot_index):
    """Write demo `InstallRecord`s onto the row for their person and title.

    `person_index` maps a name to its 0-based position on People, `slot_index` a title to its position on Catalog.
    """
    for record in records:
        row = inventory_row(person_index[record.person], slot_index[record.software])
        ws[f"{Inv.STATE}{row}"] = record.state
        ws[f"{Inv.VERSION}{row}"] = number_if_plain(record.version) or None
        ws[f"{Inv.INSTALLED_ON}{row}"] = record.installed_on
        ws[f"{Inv.ASSIGNMENT}{row}"] = record.assignment or None
        ws[f"{Inv.APPL_OVERRIDE}{row}"] = record.applicability_override or None
        ws[f"{Inv.VERIFIED_ON}{row}"] = record.verified_on
        ws[f"{Inv.VERIFIED_BY}{row}"] = record.verified_by or None
        ws[f"{Inv.NOTES}{row}"] = record.notes or None
