"""Role Defaults sheet: a role x title grid saying what each role must have, may have, or does not need."""
import validation
from layout import FIRST_ROW, HEADER_ROW, N_TITLES, ROLE_COL, ROLE_GRID_LAST_COL, ROLES_LAST, Cat, Sheet, col_letter
from starter_data import ROLES, TITLE_NAMES, role_default
from styles import (
    TAB_REFERENCE,
    Col,
    applicability_colours,
    clear_when_blank,
    header_cell,
    page_setup,
    protect_sheet,
    style_cells,
    title_block,
    unlock,
    write_headers,
)

ROLE_COL_WIDTH = 26
TITLE_COL_WIDTH = 13
HEADER_HEIGHT = 60
GRID = f"B{FIRST_ROW}:{ROLE_GRID_LAST_COL}{ROLES_LAST}"


def _write_title_headers(ws):
    """Header cells that mirror the Catalog's title order, so the grid stays aligned with it."""
    for slot in range(N_TITLES):
        catalog_cell = f"Catalog!${Cat.NAME}${FIRST_ROW + slot}"
        header_cell(ws, HEADER_ROW, 2 + slot, f'=IF({catalog_cell}="","",{catalog_cell})',
                    calc=True, width=TITLE_COL_WIDTH)
    ws.row_dimensions[HEADER_ROW].height = HEADER_HEIGHT


def create(wb):
    ws = wb.create_sheet(Sheet.ROLES)
    title_block(ws, "Role Defaults", "What each role is expected to have. Software columns follow the Catalog order. "
                                     "Blank = Optional. Add roles in the next empty row; do not insert or delete.")
    ws.sheet_properties.tabColor = TAB_REFERENCE
    write_headers(ws, [Col(ROLE_COL, "Role", ROLE_COL_WIDTH)], HEADER_HEIGHT)
    _write_title_headers(ws)
    style_cells(ws, f"{ROLE_COL}{FIRST_ROW}:{ROLE_COL}{ROLES_LAST}")
    style_cells(ws, GRID, align="center")
    unlock(ws, f"{ROLE_COL}{FIRST_ROW}:{ROLE_COL}{ROLES_LAST}")
    unlock(ws, GRID)
    validation.list_rule(ws, GRID, "=ApplicabilityList", "Applicability", "Required, Optional or Not Applicable.")
    applicability_colours(ws, GRID)
    clear_when_blank(ws, f"B{HEADER_ROW}:{ROLE_GRID_LAST_COL}{HEADER_ROW}", f"B${HEADER_ROW}")
    ws.freeze_panes = f"B{FIRST_ROW}"
    page_setup(ws, f"{HEADER_ROW}:{HEADER_ROW}")
    protect_sheet(ws)


def fill_starter(ws):
    """Starter roles with a full Required / Optional / Not Applicable row each."""
    for role_offset, role in enumerate(ROLES):
        row = FIRST_ROW + role_offset
        ws[f"{ROLE_COL}{row}"] = role
        for slot, title in enumerate(TITLE_NAMES):
            ws[f"{col_letter(2 + slot)}{row}"] = role_default(role, title)
