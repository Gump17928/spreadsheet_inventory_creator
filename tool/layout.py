"""Where things are: sheet names, capacities and column letters.

Everything that two sheets must agree on (which column holds the status, how many rows a table has) lives here.
"""
from openpyxl.utils import column_index_from_string, get_column_letter

HEADER_ROW = 5
FIRST_ROW = 6

N_PEOPLE = 100
N_TITLES = 30
N_ROLES = 30
N_RECORDS = N_PEOPLE * N_TITLES    # Inventory has one row for every person and every title slot
N_LIST_SLOTS = 20       # rows in each editable list on the Lists sheet
N_SUMMARY_ROWS = 12     # rows in the Dashboard group and department tables

PEOPLE_LAST = FIRST_ROW + N_PEOPLE - 1
CATALOG_LAST = FIRST_ROW + N_TITLES - 1
ROLES_LAST = FIRST_ROW + N_ROLES - 1
INV_LAST = FIRST_ROW + N_RECORDS - 1
LIST_LAST = FIRST_ROW + N_LIST_SLOTS - 1

WARN_DAYS = "WarnDays"    # defined names for the two thresholds on the Lists sheet
STALE_DAYS = "StaleDays"


def col_letter(index):
    """1-based column number to its letter."""
    return get_column_letter(index)


def col_index(letter):
    """Column letter to its 1-based number."""
    return column_index_from_string(letter)


def inventory_row(person_index, slot):
    """Inventory row for the person at `person_index` on People and the title at `slot` on the Catalog (both 0-based).

    Row position is the only link between an Inventory row and its person and title, so these never move.
    """
    return FIRST_ROW + person_index * N_TITLES + slot


def people_row_for(inv_row):
    """People row of the person an Inventory row belongs to."""
    return FIRST_ROW + (inv_row - FIRST_ROW) // N_TITLES


def slot_for(inv_row):
    """0-based Catalog slot of the title an Inventory row belongs to."""
    return (inv_row - FIRST_ROW) % N_TITLES


def catalog_row_for(inv_row):
    """Catalog row of the title an Inventory row belongs to."""
    return FIRST_ROW + slot_for(inv_row)


class Sheet:
    START = "About"
    LEGEND = "Key"
    INVENTORY = "Inventory"
    MATRIX = "Matrix"
    PERSON_VIEW = "Person View"
    DASHBOARD = "Dashboard"
    CATALOG = "Catalog"
    PEOPLE = "People"
    ROLES = "Role Defaults"
    LISTS = "Lists"


SHEET_ORDER = [Sheet.START, Sheet.LEGEND, Sheet.INVENTORY, Sheet.MATRIX, Sheet.PERSON_VIEW,
               Sheet.DASHBOARD, Sheet.CATALOG, Sheet.PEOPLE, Sheet.ROLES, Sheet.LISTS]


class Inv:
    """Inventory columns."""
    PERSON, COMPUTER, SOFTWARE, GROUP, STATE, VERSION = "A", "B", "C", "D", "E", "F"
    INSTALLED_ON, ASSIGNMENT, EXPIRY_OVERRIDE, APPL_OVERRIDE = "G", "H", "I", "J"
    VERIFIED_ON, VERIFIED_BY, NOTES = "K", "L", "M"
    APPLICABILITY, EXPIRY, VERSION_CHECK, LICENSE_CHECK, DAYS_SINCE, STATUS, LIVE = "N", "O", "P", "Q", "R", "S", "T"
    LAST_CALC = "T"


class Cat:
    """Catalog columns."""
    NAME, GROUP, CATEGORY, VENDOR, APPROVED, MODEL, SERVER, SEATS, EXPIRY, ADMIN, NOTES = "ABCDEFGHIJK"
    INSTALLS, SEATS_LEFT, SEAT_STATUS, DAYS_LEFT, LICENSE_STATUS = "LMNOP"
    LAST_INPUT, FIRST_CALC, LAST_CALC = "K", "L", "P"


class Ppl:
    """People columns."""
    NAME, EMPLOYEE_ID, DEPARTMENT, ROLE, ACTIVE, LOCATION, COMPUTER, MANAGER, NOTES = "ABCDEFGHI"


class Lst:
    """Lists columns and setting cells."""
    INSTALL_STATE, APPLICABILITY, LICENSE_ASSIGN, LICENSE_MODEL, GROUP, CATEGORY, DEPARTMENT = "ABCDEFG"
    SETTING_LABEL, SETTING_VALUE = "I", "J"
    WARN_CELL, STALE_CELL = "J6", "J7"


class Mat:
    """Matrix columns; software columns start at FIRST_TITLE_COL."""
    PERSON, DEPARTMENT, ROLE, COMPUTER, HAVE, MISSING, ATTENTION = "ABCDEFG"
    FIRST_TITLE_COL = 8
    GROUP_ROW = 4

    @staticmethod
    def title_col(slot):
        """Letter of the column for catalog slot `slot` (0-based)."""
        return col_letter(Mat.FIRST_TITLE_COL + slot)


MATRIX_FIRST_TITLE = Mat.title_col(0)
MATRIX_LAST_TITLE = Mat.title_col(N_TITLES - 1)
ROLE_GRID_LAST_COL = col_letter(1 + N_TITLES)   # Role Defaults: column A is the role, then one column per title
ROLE_COL = "A"
