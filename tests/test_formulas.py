"""Formula text, row mapping and workbook structure. No Excel needed."""
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tool"))

from openpyxl import load_workbook

import build_workbook
import inventory_formulas
import lookups
import sheet_matrix
from layout import (
    FIRST_ROW,
    INV_LAST,
    N_PEOPLE,
    N_RECORDS,
    N_TITLES,
    SHEET_ORDER,
    Inv,
    Sheet,
    catalog_row_for,
    inventory_row,
    people_row_for,
    slot_for,
)
from vocab import Status

EXCEL_FORMULA_LIMIT = 8192     # characters
EXCEL_NESTING_LIMIT = 64
SAMPLE_ROW = FIRST_ROW


def nesting_depth(formula):
    depth = deepest = 0
    in_text = False
    for char in formula:
        if char == '"':
            in_text = not in_text
        elif not in_text and char == "(":
            depth += 1
            deepest = max(deepest, depth)
        elif not in_text and char == ")":
            depth -= 1
    return deepest, depth


class RowMappingTests(unittest.TestCase):
    def test_inventory_has_one_row_per_person_and_title_slot(self):
        self.assertEqual(N_RECORDS, N_PEOPLE * N_TITLES)
        self.assertEqual(INV_LAST, FIRST_ROW + N_RECORDS - 1)

    def test_mapping_round_trips(self):
        for person_index in (0, 1, N_PEOPLE - 1):
            for slot in (0, 1, N_TITLES - 1):
                row = inventory_row(person_index, slot)
                self.assertEqual(people_row_for(row), FIRST_ROW + person_index)
                self.assertEqual(slot_for(row), slot)
                self.assertEqual(catalog_row_for(row), FIRST_ROW + slot)

    def test_first_and_last_rows_are_the_corners(self):
        self.assertEqual(inventory_row(0, 0), FIRST_ROW)
        self.assertEqual(inventory_row(N_PEOPLE - 1, N_TITLES - 1), INV_LAST)


class FormulaTextTests(unittest.TestCase):
    def inventory_formulas(self):
        builders = [inventory_formulas.person, inventory_formulas.software, inventory_formulas.computer,
                    inventory_formulas.group, inventory_formulas.live, inventory_formulas.applicability,
                    inventory_formulas.expiry, inventory_formulas.version_check, inventory_formulas.license_check,
                    inventory_formulas.days_since_verified, inventory_formulas.status]
        return [build(SAMPLE_ROW) for build in builders]

    def test_formulas_are_balanced_and_within_excel_limits(self):
        formulas = self.inventory_formulas() + [sheet_matrix.cell_formula(SAMPLE_ROW, slot) for slot in range(N_TITLES)]
        for formula in formulas:
            deepest, remaining = nesting_depth(formula)
            self.assertEqual(remaining, 0, formula)
            self.assertLessEqual(deepest, EXCEL_NESTING_LIMIT, formula)
            self.assertLessEqual(len(formula), EXCEL_FORMULA_LIMIT, formula)
            self.assertTrue(formula.startswith("="))

    def test_formulas_contain_plain_words_not_enum_names(self):
        for formula in self.inventory_formulas():
            self.assertNotIn("Status.", formula)
            self.assertNotIn("InstallState.", formula)

    def test_inventory_row_points_at_its_own_person_and_title(self):
        row = inventory_row(2, 3)
        self.assertIn(f"People!$A${FIRST_ROW + 2}", inventory_formulas.person(row))
        self.assertIn(f"Catalog!$A${FIRST_ROW + 3}", inventory_formulas.software(row))

    def test_matrix_cell_points_at_the_inventory_row_for_its_person_and_title(self):
        formula = sheet_matrix.cell_formula(FIRST_ROW + 2, 3)
        self.assertIn(f"Inventory!$S${inventory_row(2, 3)}", formula)

    def test_not_installed_status_mentions_all_three_outcomes(self):
        text = lookups.not_installed_status("$N6")
        for status in (Status.NOT_APPLICABLE, Status.MISSING, Status.NOT_INSTALLED):
            self.assertIn(f'"{status}"', text)

    def test_role_default_treats_blank_as_optional(self):
        self.assertIn('"Optional"', lookups.role_default_by_slot("$C6", 0))


class WorkbookStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.template = Path(cls.tmp.name) / "template.xlsx"
        cls.example = Path(cls.tmp.name) / "example.xlsx"
        build_workbook.build_template(cls.template)
        build_workbook.build_example(cls.example)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_sheets_are_in_the_documented_order(self):
        self.assertEqual(load_workbook(self.template).sheetnames, SHEET_ORDER)

    def test_template_has_catalog_and_roles_but_no_people(self):
        wb = load_workbook(self.template)
        self.assertEqual(wb[Sheet.CATALOG]["A6"].value, "AutoCAD")
        self.assertEqual(wb[Sheet.ROLES]["A6"].value, "Mechanical Engineer")
        self.assertIsNone(wb[Sheet.PEOPLE]["A6"].value)
        self.assertIsNone(wb[Sheet.INVENTORY][f"{Inv.STATE}6"].value)

    def test_example_has_people_and_installs(self):
        wb = load_workbook(self.example)
        self.assertEqual(wb[Sheet.PEOPLE]["A6"].value, "Alex Rivera")
        states = wb[Sheet.INVENTORY][f"{Inv.STATE}{FIRST_ROW}:{Inv.STATE}{INV_LAST}"]
        self.assertGreater(sum(1 for row in states if row[0].value == "Installed"), 0)

    def test_dropdown_names_are_defined(self):
        names = set(load_workbook(self.template).defined_names)
        expected = {"PeopleList", "RoleList", "InstallStates", "ApplicabilityList", "LicenseAssignList",
                    "LicenseModelList", "GroupList", "CategoryList", "DeptList", "WarnDays", "StaleDays"}
        self.assertLessEqual(expected, names)

    def test_inventory_dropdowns_cover_only_typed_columns(self):
        rules = load_workbook(self.template)[Sheet.INVENTORY].data_validations.dataValidation
        lists = {rule.formula1 for rule in rules if rule.type == "list"}
        self.assertLessEqual({"=InstallStates", "=LicenseAssignList", "=ApplicabilityList"}, lists)
        self.assertNotIn("=PeopleList", lists)

    def test_structure_sheets_are_protected_against_row_changes(self):
        wb = load_workbook(self.template)
        for name in (Sheet.INVENTORY, Sheet.PEOPLE, Sheet.CATALOG, Sheet.ROLES, Sheet.MATRIX, Sheet.PERSON_VIEW,
                     Sheet.DASHBOARD):
            protection = wb[name].protection
            self.assertTrue(protection.sheet, name)
            self.assertTrue(protection.insertRows and protection.deleteRows and protection.sort, name)  # True = blocked
        self.assertFalse(wb[Sheet.LISTS].protection.sheet)

    def test_typed_cells_are_unlocked_and_calculated_cells_locked(self):
        ws = load_workbook(self.template)[Sheet.INVENTORY]
        self.assertFalse(ws[f"{Inv.STATE}6"].protection.locked)
        self.assertFalse(ws[f"{Inv.NOTES}6"].protection.locked)
        self.assertTrue(ws[f"{Inv.STATUS}6"].protection.locked)
        self.assertTrue(ws[f"{Inv.PERSON}6"].protection.locked)


if __name__ == "__main__":
    unittest.main()
