"""People sheet: the roster. Role decides each person's default software applicability."""
import validation
from demo_data import BUSINESS_DEPARTMENT, LOCATION, MANAGER_BUSINESS, MANAGER_ENGINEERING, PEOPLE
from dynamic_names import ROLE_LIST
from layout import FIRST_ROW, HEADER_ROW, PEOPLE_LAST, Ppl, Sheet
from styles import TAB_REFERENCE, Col, page_setup, protect_sheet, style_columns, title_block, write_headers
from vocab import Active, values

COLS = [
    Col(Ppl.NAME, "Name", 24), Col(Ppl.EMPLOYEE_ID, "Employee ID", 16), Col(Ppl.DEPARTMENT, "Department", 18),
    Col(Ppl.ROLE, "Role", 24), Col(Ppl.ACTIVE, "Active", 10, align="center"), Col(Ppl.LOCATION, "Location", 16),
    Col(Ppl.COMPUTER, "Computer / Asset Tag", 26), Col(Ppl.MANAGER, "Manager", 20), Col(Ppl.NOTES, "Notes", 30),
]
FIRST_EMPLOYEE_ID = 1001
FIRST_ASSET_NUMBER = 101
ENGINEERING_PREFIX, BUSINESS_PREFIX = "ENG", "BUS"
ACTIVE_CHOICES = '"' + ",".join(values(Active)) + '"'


def _cells(col):
    return f"{col}{FIRST_ROW}:{col}{PEOPLE_LAST}"


def create(wb):
    ws = wb.create_sheet(Sheet.PEOPLE)
    title_block(ws, "People", "One row per person. Role drives each person's Required / Optional / Not Applicable "
                              "software. Add people in the next empty row - never insert, delete or sort rows.")
    ws.sheet_properties.tabColor = TAB_REFERENCE
    write_headers(ws, COLS, height=36)
    style_columns(ws, COLS, PEOPLE_LAST)
    validation.list_rule(ws, _cells(Ppl.DEPARTMENT), "=DeptList", "Department", "Pick a department.")
    validation.list_rule(ws, _cells(Ppl.ROLE), f"={ROLE_LIST}", "Role",
                         "Role decides default software. Edit roles on the Role Defaults sheet.")
    validation.list_rule(ws, _cells(Ppl.ACTIVE), ACTIVE_CHOICES, "Active",
                         "Leave blank or Yes while they are here. Set No when they leave - do not delete the row.")
    validation.unique_rule(ws, _cells(Ppl.NAME), f"${Ppl.NAME}${FIRST_ROW}:${Ppl.NAME}${PEOPLE_LAST}",
                           f"{Ppl.NAME}{FIRST_ROW}", "name")
    ws.freeze_panes = f"B{FIRST_ROW}"
    ws.auto_filter.ref = f"{Ppl.NAME}{HEADER_ROW}:{Ppl.NOTES}{PEOPLE_LAST}"
    page_setup(ws, f"{HEADER_ROW}:{HEADER_ROW}")
    protect_sheet(ws)


def fill_demo(ws):
    for offset, person in enumerate(PEOPLE):
        row = FIRST_ROW + offset
        business = person.department == BUSINESS_DEPARTMENT
        prefix = BUSINESS_PREFIX if business else ENGINEERING_PREFIX
        ws[f"{Ppl.NAME}{row}"] = person.name
        ws[f"{Ppl.EMPLOYEE_ID}{row}"] = f"E{FIRST_EMPLOYEE_ID + offset}"
        ws[f"{Ppl.DEPARTMENT}{row}"] = person.department
        ws[f"{Ppl.ROLE}{row}"] = person.role
        ws[f"{Ppl.ACTIVE}{row}"] = Active.YES
        ws[f"{Ppl.LOCATION}{row}"] = LOCATION
        ws[f"{Ppl.COMPUTER}{row}"] = f"{prefix}-LT-{FIRST_ASSET_NUMBER + offset:04d}"
        ws[f"{Ppl.MANAGER}{row}"] = MANAGER_BUSINESS if business else MANAGER_ENGINEERING
