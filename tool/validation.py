"""Excel data-validation rules (dropdowns, dates, numbers, uniqueness) with helpful prompts."""
from openpyxl.worksheet.datavalidation import DataValidation

MIN_DATE = "DATE(2000,1,1)"
MAX_DATE = "DATE(2100,12,31)"
MAX_VERSION_LENGTH = 30
MAX_DAYS_SETTING = 3650


def _add(ws, cells, rule):
    ws.add_data_validation(rule)
    rule.add(cells)


def list_rule(ws, cells, source, title, prompt):
    """Dropdown. `source` is a formula such as '=PeopleList'."""
    _add(ws, cells, DataValidation(
        type="list", formula1=source, allow_blank=True,
        showErrorMessage=True, errorTitle="Invalid entry", error="Pick a value from the dropdown list.",
        showInputMessage=True, promptTitle=title, prompt=prompt))


def date_rule(ws, cells, title, prompt, not_future=False):
    """Real dates only; optionally not later than today."""
    upper = "TODAY()" if not_future else MAX_DATE
    suffix = ", not in the future." if not_future else "."
    _add(ws, cells, DataValidation(
        type="date", operator="between", formula1=MIN_DATE, formula2=upper, allow_blank=True,
        showErrorMessage=True, errorTitle="Invalid date", error="Enter a real date (e.g. 2026-11-12)" + suffix,
        showInputMessage=True, promptTitle=title, prompt=prompt))


def count_rule(ws, cells, title, prompt):
    """Whole number, zero or more."""
    _add(ws, cells, DataValidation(
        type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
        showErrorMessage=True, errorTitle="Invalid number", error="Enter a whole number, 0 or more.",
        showInputMessage=True, promptTitle=title, prompt=prompt))


def days_rule(ws, cells):
    _add(ws, cells, DataValidation(
        type="whole", operator="between", formula1="1", formula2=str(MAX_DAYS_SETTING),
        showErrorMessage=True, errorTitle="Invalid", error=f"Enter a whole number of days (1-{MAX_DAYS_SETTING})."))


def short_text_rule(ws, cells, title, prompt):
    _add(ws, cells, DataValidation(
        type="textLength", operator="lessThanOrEqual", formula1=str(MAX_VERSION_LENGTH), allow_blank=True,
        showErrorMessage=True, errorTitle="Too long", error=f"Keep this to {MAX_VERSION_LENGTH} characters or fewer.",
        showInputMessage=True, promptTitle=title, prompt=prompt))


def unique_rule(ws, cells, column_range, first_cell, what):
    """Reject a value that already exists in `column_range`."""
    _add(ws, cells, DataValidation(
        type="custom", formula1=f"COUNTIF({column_range},{first_cell})=1",
        showErrorMessage=True, errorTitle="Duplicate", error=f"That {what} already exists. Names must be unique."))
