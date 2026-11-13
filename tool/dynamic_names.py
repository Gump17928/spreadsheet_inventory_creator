"""Defined names for the dropdowns that follow data on other sheets (people, roles, software)."""
from openpyxl.workbook.defined_name import DefinedName

import refs
from layout import PEOPLE_LAST, ROLE_COL, ROLES_LAST, Ppl, Sheet

PEOPLE_LIST = "PeopleList"
ROLE_LIST = "RoleList"

GROWING_LISTS = {
    PEOPLE_LIST: (Sheet.PEOPLE, Ppl.NAME, PEOPLE_LAST),
    ROLE_LIST: (Sheet.ROLES, ROLE_COL, ROLES_LAST),
}


def define_growing_lists(wb):
    for name, (sheet, col, last_row) in GROWING_LISTS.items():
        wb.defined_names[name] = DefinedName(name, attr_text=refs.growing_list(sheet, col, last_row))
