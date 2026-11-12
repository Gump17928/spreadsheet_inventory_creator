"""Defined names for the dropdowns that follow data on other sheets (people, roles, software)."""
from openpyxl.workbook.defined_name import DefinedName

import refs
from layout import CATALOG_LAST, PEOPLE_LAST, ROLE_COL, ROLES_LAST, Cat, Ppl, Sheet

PEOPLE_LIST = "PeopleList"
ROLE_LIST = "RoleList"
SOFTWARE_LIST = "SoftwareList"

GROWING_LISTS = {
    PEOPLE_LIST: (Sheet.PEOPLE, Ppl.NAME, PEOPLE_LAST),
    ROLE_LIST: (Sheet.ROLES, ROLE_COL, ROLES_LAST),
    SOFTWARE_LIST: (Sheet.CATALOG, Cat.NAME, CATALOG_LAST),
}


def define_growing_lists(wb):
    for name, (sheet, col, last_row) in GROWING_LISTS.items():
        wb.defined_names[name] = DefinedName(name, attr_text=refs.growing_list(sheet, col, last_row))
