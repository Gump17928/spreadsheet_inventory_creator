"""Formula text and workbook structure. No Excel needed."""
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
from layout import FIRST_ROW, N_TITLES, SHEET_ORDER, Inv, Sheet
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


class FormulaTextTests(unittest.TestCase):
    def inventory_formulas(self):
        builders = [inventory_formulas.key, inventory_formulas.computer, inventory_formulas.group,
                    inventory_formulas.applicability, inventory_formulas.expiry, inventory_formulas.version_check,
                    inventory_formulas.license_check, inventory_formulas.days_since_verified, inventory_formulas.status]
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

    def test_template_has_catalog_and_roles_but_no_people_or_installs(self):
        wb = load_workbook(self.template)
        self.assertEqual(wb[Sheet.CATALOG]["A6"].value, "AutoCAD")
        self.assertEqual(wb[Sheet.ROLES]["A6"].value, "Mechanical Engineer")
        self.assertIsNone(wb[Sheet.PEOPLE]["A6"].value)
        self.assertIsNone(wb[Sheet.INVENTORY][f"{Inv.PERSON}6"].value)

    def test_example_has_people_and_installs(self):
        wb = load_workbook(self.example)
        self.assertEqual(wb[Sheet.PEOPLE]["A6"].value, "Alex Rivera")
        self.assertIsNotNone(wb[Sheet.INVENTORY][f"{Inv.PERSON}6"].value)

    def test_dropdown_names_are_defined(self):
        names = set(load_workbook(self.template).defined_names)
        expected = {"PeopleList", "RoleList", "SoftwareList", "InstallStates", "ApplicabilityList", "LicenseAssignList",
                    "LicenseModelList", "GroupList", "CategoryList", "DeptList", "WarnDays", "StaleDays"}
        self.assertLessEqual(expected, names)

    def test_inventory_has_dropdowns_on_the_input_columns(self):
        rules = load_workbook(self.template)[Sheet.INVENTORY].data_validations.dataValidation
        lists = {rule.formula1 for rule in rules if rule.type == "list"}
        self.assertLessEqual({"=PeopleList", "=SoftwareList", "=InstallStates"}, lists)

    def test_report_sheets_are_protected_and_inventory_is_not(self):
        wb = load_workbook(self.template)
        for name in (Sheet.MATRIX, Sheet.PERSON_VIEW, Sheet.DASHBOARD):
            self.assertTrue(wb[name].protection.sheet, name)
        self.assertFalse(wb[Sheet.INVENTORY].protection.sheet)


if __name__ == "__main__":
    unittest.main()
