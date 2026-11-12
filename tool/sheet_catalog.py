"""Catalog sheet: one row per software title with license, seat and expiry tracking."""
from datetime import date, timedelta

import validation
from cell_values import number_if_plain
from demo_data import LICENSE_DETAILS
from layout import CATALOG_LAST, FIRST_ROW, HEADER_ROW, WARN_DAYS, Cat, Inv, Sheet
from refs import inv
from starter_data import TITLES
from styles import (
    TAB_REFERENCE,
    Col,
    band,
    date_flags,
    note,
    page_setup,
    style_columns,
    text_flags,
    title_block,
    write_headers,
)
from vocab import InstallState, LicenseCheck, LicenseModel, SeatStatus

COLS = [
    Col(Cat.NAME, "Software", 26), Col(Cat.GROUP, "Group / Suite", 20), Col(Cat.CATEGORY, "Category", 24),
    Col(Cat.VENDOR, "Vendor", 18), Col(Cat.APPROVED, "Approved Version", 14, align="center"),
    Col(Cat.MODEL, "License Model", 26), Col(Cat.SERVER, "License Server (host:port)", 26),
    Col(Cat.SEATS, "Seats Purchased", 13, align="center", fmt="0"),
    Col(Cat.EXPIRY, "License / Contract Expiry", 17, align="center", fmt="yyyy-mm-dd"),
    Col(Cat.ADMIN, "License Admin", 16), Col(Cat.NOTES, "Notes", 28),
    Col(Cat.INSTALLS, "Installs Recorded", 13, True, "0", "center"),
    Col(Cat.SEATS_LEFT, "Seats Remaining", 13, True, "0", "center"),
    Col(Cat.SEAT_STATUS, "Seat Status", 26, True, align="center"),
    Col(Cat.DAYS_LEFT, "Days to Expiry", 12, True, "0", "center"),
    Col(Cat.LICENSE_STATUS, "License Status", 16, True, align="center"),
]
INPUT_RANGE = f"{Cat.NAME}{FIRST_ROW}:{Cat.NOTES}{CATALOG_LAST}"
SEAT_STATUS_FLAGS = {SeatStatus.OVER: "9C0006", SeatStatus.MONITOR: "7F6000"}
LICENSE_STATUS_FLAGS = {LicenseCheck.EXPIRED: "C00000", LicenseCheck.EXPIRING: "9C4A00"}


def _cell(col, row):
    return f"${col}{row}"


def _installs(row):
    name = _cell(Cat.NAME, row)
    return (f'=IF({name}="","",COUNTIFS({inv(Inv.SOFTWARE)},{name},'
            f'{inv(Inv.STATE)},"{InstallState.INSTALLED}"))')


def _seats_left(row):
    name, seats = _cell(Cat.NAME, row), _cell(Cat.SEATS, row)
    return f'=IF(OR({name}="",{seats}=""),"",{seats}-{_cell(Cat.INSTALLS, row)})'


def _seat_status(row):
    name, seats = _cell(Cat.NAME, row), _cell(Cat.SEATS, row)
    installs, model = _cell(Cat.INSTALLS, row), _cell(Cat.MODEL, row)
    over = f'IF({model}="{LicenseModel.FLOATING}","{SeatStatus.MONITOR}","{SeatStatus.OVER}")'
    return (f'=IF({name}="","",IF({seats}="","{SeatStatus.NOT_TRACKED}",'
            f'IF({installs}>{seats},{over},"{SeatStatus.OK}")))')


def _days_left(row):
    name, expiry = _cell(Cat.NAME, row), _cell(Cat.EXPIRY, row)
    return f'=IF(OR({name}="",{expiry}=""),"",{expiry}-TODAY())'


def _license_status(row):
    name, expiry = _cell(Cat.NAME, row), _cell(Cat.EXPIRY, row)
    return (f'=IF({name}="","",IF({expiry}="","{LicenseCheck.NO_EXPIRY}",IF({expiry}<TODAY(),"{LicenseCheck.EXPIRED}",'
            f'IF({expiry}<=TODAY()+{WARN_DAYS},"{LicenseCheck.EXPIRING}","{LicenseCheck.VALID}"))))')


FORMULAS = {Cat.INSTALLS: _installs, Cat.SEATS_LEFT: _seats_left, Cat.SEAT_STATUS: _seat_status,
            Cat.DAYS_LEFT: _days_left, Cat.LICENSE_STATUS: _license_status}


def _write_formulas(ws):
    for row in range(FIRST_ROW, CATALOG_LAST + 1):
        for letter, build in FORMULAS.items():
            ws[f"{letter}{row}"] = build(row)


def _add_validation(ws):
    cells = lambda col: f"{col}{FIRST_ROW}:{col}{CATALOG_LAST}"  # noqa: E731
    validation.list_rule(ws, cells(Cat.GROUP), "=GroupList", "Group / Suite", "Product family, e.g. Autodesk.")
    validation.list_rule(ws, cells(Cat.CATEGORY), "=CategoryList", "Category", "What kind of software this is.")
    validation.list_rule(ws, cells(Cat.MODEL), "=LicenseModelList", "License Model", "How this title is licensed.")
    validation.date_rule(ws, cells(Cat.EXPIRY), "Expiry", "License/contract end date. Blank for perpetual or free.")
    validation.count_rule(ws, cells(Cat.SEATS), "Seats", "Total seats or licenses purchased.")
    validation.unique_rule(ws, cells(Cat.NAME), f"${Cat.NAME}${FIRST_ROW}:${Cat.NAME}${CATALOG_LAST}",
                           f"{Cat.NAME}{FIRST_ROW}", "software name")


def _add_formatting(ws):
    band(ws, INPUT_RANGE)
    text_flags(ws, f"{Cat.SEAT_STATUS}{FIRST_ROW}:{Cat.SEAT_STATUS}{CATALOG_LAST}", SEAT_STATUS_FLAGS)
    text_flags(ws, f"{Cat.LICENSE_STATUS}{FIRST_ROW}:{Cat.LICENSE_STATUS}{CATALOG_LAST}", LICENSE_STATUS_FLAGS)
    date_flags(ws, f"{Cat.EXPIRY}{FIRST_ROW}:{Cat.EXPIRY}{CATALOG_LAST}")


def create(wb):
    ws = wb.create_sheet(Sheet.CATALOG)
    title_block(ws, "Software Catalog", "One row per software title. Defines groups, approved versions and license "
                                        "details. Add new titles in the next empty row.")
    ws.sheet_properties.tabColor = TAB_REFERENCE
    note(ws, "A3", "Navy headers = you type or choose.  Grey headers and cells = calculated automatically.")
    write_headers(ws, COLS)
    style_columns(ws, COLS, CATALOG_LAST)
    _write_formulas(ws)
    _add_validation(ws)
    _add_formatting(ws)
    ws.freeze_panes = f"B{FIRST_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{Cat.LAST_CALC}{CATALOG_LAST}"
    page_setup(ws, f"{HEADER_ROW}:{HEADER_ROW}")


def fill_starter(ws):
    """Titles, groups and approved versions only - no license numbers."""
    for offset, title in enumerate(TITLES):
        row = FIRST_ROW + offset
        ws[f"{Cat.NAME}{row}"] = title.name
        ws[f"{Cat.GROUP}{row}"] = title.group
        ws[f"{Cat.CATEGORY}{row}"] = title.category
        ws[f"{Cat.VENDOR}{row}"] = title.vendor
        ws[f"{Cat.APPROVED}{row}"] = number_if_plain(title.approved_version) or None
        ws[f"{Cat.MODEL}{row}"] = title.model


def fill_demo(ws, today=None):
    """Fake server names, seat counts and expiry dates for the example workbook."""
    today = today or date.today()
    for offset, title in enumerate(TITLES):
        detail = LICENSE_DETAILS[title.name]
        row = FIRST_ROW + offset
        ws[f"{Cat.SERVER}{row}"] = detail.server or None
        ws[f"{Cat.SEATS}{row}"] = detail.seats
        ws[f"{Cat.ADMIN}{row}"] = detail.admin
        if detail.expires_in_days is not None:
            ws[f"{Cat.EXPIRY}{row}"] = today + timedelta(days=detail.expires_in_days)
