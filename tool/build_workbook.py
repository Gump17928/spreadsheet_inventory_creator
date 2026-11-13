"""Driver: builds the blank template and the example workbook.

Run from the project folder:  .venv\\Scripts\\python.exe tool\build_workbook.py
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook

import demo_data
import demo_inventory
import dynamic_names
import sheet_catalog
import sheet_dashboard
import sheet_inventory
import sheet_legend
import sheet_lists
import sheet_matrix
import sheet_people
import sheet_person_view
import sheet_roles
import sheet_start
from layout import SHEET_ORDER, Sheet

PROJECT_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_FILE = "Software_Inventory.xlsx"
EXAMPLE_FILE = "Software_Inventory_Example.xlsx"


def create_structure():
    """All sheets, formulas, dropdowns and formatting - no people or installs."""
    wb = Workbook()
    sheet_start.fill_start(wb.active)
    sheet_legend.create(wb)
    sheet_lists.create(wb)
    sheet_catalog.create(wb)
    sheet_people.create(wb)
    sheet_roles.create(wb)
    sheet_inventory.create(wb)
    sheet_matrix.create(wb)
    sheet_person_view.create(wb)
    sheet_dashboard.create(wb)
    dynamic_names.define_growing_lists(wb)
    wb._sheets = [wb[name] for name in SHEET_ORDER]   # openpyxl has no public "reorder sheets"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True              # Excel computes every formula when the file is opened
    return wb


def add_starter_content(wb):
    """Sample catalog titles and role defaults, so the blank template is not an empty shell."""
    sheet_catalog.fill_starter(wb[Sheet.CATALOG])
    sheet_roles.fill_starter(wb[Sheet.ROLES])


def add_demo_content(wb, today=None):
    """Fake license numbers, people and installs for the example workbook."""
    today = today or date.today()
    sheet_catalog.fill_demo(wb[Sheet.CATALOG], today)
    sheet_people.fill_demo(wb[Sheet.PEOPLE])
    sheet_inventory.write_records(wb[Sheet.INVENTORY], demo_inventory.make_records(today),
                                  demo_inventory.PERSON_INDEX, demo_inventory.SLOT_INDEX)
    sheet_person_view.select_person(wb[Sheet.PERSON_VIEW], demo_data.PEOPLE[0].name)


def build_template(path):
    wb = create_structure()
    add_starter_content(wb)
    wb.save(path)


def build_example(path):
    wb = create_structure()
    add_starter_content(wb)
    add_demo_content(wb)
    wb.save(path)


def main():
    build_template(PROJECT_DIR / TEMPLATE_FILE)
    build_example(PROJECT_DIR / EXAMPLE_FILE)
    print(f"Wrote {TEMPLATE_FILE} and {EXAMPLE_FILE} in {PROJECT_DIR}")


if __name__ == "__main__":
    main()
