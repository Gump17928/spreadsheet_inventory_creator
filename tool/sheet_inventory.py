"""Inventory sheet: one row per person + software. The only sheet people type into day to day."""
import inventory_formulas as formulas
import validation
from cell_values import number_if_plain
from dynamic_names import PEOPLE_LIST, SOFTWARE_LIST
from layout import FIRST_ROW, HEADER_ROW, INV_LAST, Inv, Sheet
from styles import (
    TAB_INPUT,
    Col,
    band,
    date_flags,
    note,
    page_setup,
    stale_flag,
    status_colours,
    style_columns,
    text_flags,
    title_block,
    write_headers,
)

DATE_FORMAT = "yyyy-mm-dd"
COLS = [
    Col(Inv.PERSON, "Person", 22), Col(Inv.COMPUTER, "Computer / Asset Tag", 17, True),
    Col(Inv.SOFTWARE, "Software", 24), Col(Inv.GROUP, "Group / Suite", 18, True),
    Col(Inv.STATE, "Install State", 14), Col(Inv.VERSION, "Installed Version", 12, align="center"),
    Col(Inv.INSTALLED_ON, "Install Date", 12, fmt=DATE_FORMAT, align="center"),
    Col(Inv.ASSIGNMENT, "License Assignment", 20),
    Col(Inv.EXPIRY_OVERRIDE, "License Expiry (override)", 14, fmt=DATE_FORMAT, align="center"),
    Col(Inv.APPL_OVERRIDE, "Applicability Override", 15),
    Col(Inv.VERIFIED_ON, "Last Verified", 12, fmt=DATE_FORMAT, align="center"), Col(Inv.VERIFIED_BY, "Verified By", 14),
    Col(Inv.NOTES, "Notes", 30),
    Col(Inv.APPLICABILITY, "Applicability (effective)", 15, True, align="center"),
    Col(Inv.EXPIRY, "License Expiry (effective)", 14, True, DATE_FORMAT, "center"),
    Col(Inv.VERSION_CHECK, "Version Check", 14, True, align="center"),
    Col(Inv.LICENSE_CHECK, "License Check", 14, True, align="center"),
    Col(Inv.DAYS_SINCE, "Days Since Verified", 13, True, "0", "center"),
    Col(Inv.STATUS, "STATUS", 14, True, align="center", bold=True),
    Col(Inv.KEY, "Key", 6, True),
]
FORMULAS = {
    Inv.COMPUTER: formulas.computer, Inv.GROUP: formulas.group, Inv.APPLICABILITY: formulas.applicability,
    Inv.EXPIRY: formulas.expiry, Inv.VERSION_CHECK: formulas.version_check, Inv.LICENSE_CHECK: formulas.license_check,
    Inv.DAYS_SINCE: formulas.days_since_verified, Inv.STATUS: formulas.status, Inv.KEY: formulas.key,
}
VERSION_FLAGS = {"Outdated": "9C4A00", "Check": "9C4A00", "No version": "9C4A00"}
LICENSE_FLAGS = {"Expired": "C00000", "Unlicensed": "C00000", "Expiring": "9C4A00", "Unverified": "9C4A00"}


def _cells(col):
    return f"{col}{FIRST_ROW}:{col}{INV_LAST}"


def _write_formulas(ws):
    for row in range(FIRST_ROW, INV_LAST + 1):
        for letter, build in FORMULAS.items():
            ws[f"{letter}{row}"] = build(row)


def _add_validation(ws):
    validation.list_rule(ws, _cells(Inv.PERSON), f"={PEOPLE_LIST}", "Person", "Pick from the People sheet.")
    validation.list_rule(ws, _cells(Inv.SOFTWARE), f"={SOFTWARE_LIST}", "Software", "Pick from the Catalog.")
    validation.list_rule(ws, _cells(Inv.STATE), "=InstallStates", "Install State",
                         "Installed / Not Installed / Pending (requested) / Unknown (needs a check).")
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
    band(ws, f"{Inv.FIRST_INPUT}{FIRST_ROW}:{Inv.LAST_INPUT}{INV_LAST}")
    status_colours(ws, _cells(Inv.STATUS))
    date_flags(ws, _cells(Inv.EXPIRY))
    stale_flag(ws, _cells(Inv.DAYS_SINCE))
    text_flags(ws, _cells(Inv.VERSION_CHECK), VERSION_FLAGS)
    text_flags(ws, _cells(Inv.LICENSE_CHECK), LICENSE_FLAGS)


def create(wb):
    ws = wb.create_sheet(Sheet.INVENTORY)
    title_block(ws, "Inventory", "One row per person + software that is installed, pending, or needs a note. "
                                 "No row = not installed (role defaults decide Missing vs N/A).")
    ws.sheet_properties.tabColor = TAB_INPUT
    note(ws, "A3", "Navy headers = you enter.  Grey = automatic.  Applicability Override: leave blank to use the "
                   "role default.")
    write_headers(ws, COLS)
    style_columns(ws, COLS, INV_LAST)
    _write_formulas(ws)
    _add_validation(ws)
    _add_formatting(ws)
    ws.column_dimensions[Inv.KEY].hidden = True
    ws.freeze_panes = f"{Inv.GROUP}{FIRST_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{Inv.LAST_CALC}{INV_LAST}"
    page_setup(ws, f"{HEADER_ROW}:{HEADER_ROW}")


def write_records(ws, records):
    """Write demo `InstallRecord`s from the first data row down."""
    for offset, record in enumerate(records):
        row = FIRST_ROW + offset
        ws[f"{Inv.PERSON}{row}"] = record.person
        ws[f"{Inv.SOFTWARE}{row}"] = record.software
        ws[f"{Inv.STATE}{row}"] = record.state
        ws[f"{Inv.VERSION}{row}"] = number_if_plain(record.version) or None
        ws[f"{Inv.INSTALLED_ON}{row}"] = record.installed_on
        ws[f"{Inv.ASSIGNMENT}{row}"] = record.assignment or None
        ws[f"{Inv.APPL_OVERRIDE}{row}"] = record.applicability_override or None
        ws[f"{Inv.VERIFIED_ON}{row}"] = record.verified_on
        ws[f"{Inv.VERIFIED_BY}{row}"] = record.verified_by or None
        ws[f"{Inv.NOTES}{row}"] = record.notes or None
