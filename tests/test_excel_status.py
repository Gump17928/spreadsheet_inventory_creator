"""Every status, calculated by real Excel. Skipped automatically when Excel (or pywin32) is not installed."""
import sys
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tool"))

import build_workbook
import excel_com
import sheet_inventory
import sheet_people
from demo_inventory import InstallRecord
from layout import FIRST_ROW, Inv, Mat, Sheet
from starter_data import TITLE_NAMES
from vocab import Applicability, InstallState, LicenseAssign, Status

TODAY = date.today()
ALEX, CASEY = "Alex Rivera", "Casey Morgan"      # Mechanical Engineer, Business Admin
CASEY_ROW = FIRST_ROW + 12                       # Casey's row on the People and Matrix sheets
FRESH = TODAY - timedelta(days=10)
STALE = TODAY - timedelta(days=400)
INSTALLED, NOT_INSTALLED = InstallState.INSTALLED, InstallState.NOT_INSTALLED
NAMED, FLOATING = LicenseAssign.NAMED, LicenseAssign.FLOATING


def record(software, state, version="", assignment="", verified=FRESH, person=ALEX, **extra):
    return InstallRecord(person, software, state, version, None, assignment, verified_on=verified, **extra)


# (record, expected Inventory status). Alex: AutoCAD Optional, Revit N/A, SOLIDWORKS/ANSYS Mechanical Required.
CASES = [
    (record("AutoCAD", INSTALLED, "2026", NAMED), Status.INSTALLED),
    (record("Revit", INSTALLED, "2026", NAMED), Status.REVIEW),
    (record("SOLIDWORKS", INSTALLED, "2025", FLOATING), Status.OUTDATED),
    (record("ANSYS Fluent", INSTALLED, "2024 R2", FLOATING), Status.VERIFY),             # text version differs -> Check
    (record("CATIA", INSTALLED, "2025", FLOATING, verified=STALE), Status.VERIFY),        # data is stale
    (record("MATLAB", INSTALLED, "R2025b", LicenseAssign.UNLICENSED), Status.LICENSE_ISSUE),
    (record("Microsoft 365 Apps", INSTALLED, "", LicenseAssign.SITE), Status.INSTALLED),  # no approved version tracked
    (record("Python", INSTALLED, "3.12", LicenseAssign.FREE), Status.INSTALLED),
    (record("PTC Mathcad Prime", NOT_INSTALLED), Status.MISSING),
    (record("Microsoft Visio", NOT_INSTALLED), Status.NOT_INSTALLED),
    (record("Navisworks Manage", NOT_INSTALLED), Status.NOT_APPLICABLE),
    (record("Adobe Acrobat Pro", InstallState.PENDING), Status.PENDING),
    (record("Inventor", InstallState.UNKNOWN), Status.UNKNOWN),
    (record("Bluebeam Revu", INSTALLED, "21", NAMED), Status.DUPLICATE),
    (record("Bluebeam Revu", INSTALLED, "21", NAMED), Status.DUPLICATE),
    (record("AutoCAD", NOT_INSTALLED, person=CASEY, applicability_override=Applicability.REQUIRED), Status.MISSING),
]


def matrix_address(person_row, software):
    return f"{Mat.title_col(TITLE_NAMES.index(software))}{person_row}"


@unittest.skipUnless(excel_com.excel_available(), "Excel is not available")
class ExcelStatusTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        path = Path(cls.tmp.name) / "scenarios.xlsx"
        wb = build_workbook.create_structure()
        build_workbook.add_starter_content(wb)
        sheet_people.fill_demo(wb[Sheet.PEOPLE])
        sheet_inventory.write_records(wb[Sheet.INVENTORY], [case for case, _expected in CASES])
        wb.save(path)
        with excel_com.excel_session() as app, excel_com.open_workbook(app, path) as workbook:
            cls.errors = excel_com.error_cells(workbook)
            cls.statuses = [excel_com.read_cell(workbook, Sheet.INVENTORY, f"{Inv.STATUS}{FIRST_ROW + i}")
                            for i in range(len(CASES))]
            cls.matrix = cls._read_matrix(workbook)
            cls.person_view_missing = cls._person_view_missing(workbook)

    @staticmethod
    def _read_matrix(workbook):
        """Matrix statuses for Alex's and Casey's rows, keyed by (person row, software)."""
        rows = (FIRST_ROW, CASEY_ROW)
        return {(row, software): excel_com.read_cell(workbook, Sheet.MATRIX, matrix_address(row, software))
                for row in rows for software in TITLE_NAMES}

    @staticmethod
    def _person_view_missing(workbook):
        workbook.Worksheets(Sheet.PERSON_VIEW).Range("B4").Value = ALEX
        return excel_com.read_cell(workbook, Sheet.PERSON_VIEW, "B7")

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_no_formula_errors_anywhere(self):
        self.assertEqual(self.errors, {})

    def test_every_status_is_calculated_correctly(self):
        for index, (case, expected) in enumerate(CASES):
            with self.subTest(software=case.software, person=case.person, index=index):
                self.assertEqual(self.statuses[index], expected.value)

    def test_matrix_shows_missing_for_required_title_with_no_row(self):
        self.assertEqual(self.matrix[(FIRST_ROW, "Inventor")], Status.UNKNOWN.value)   # has a row (Unknown)
        self.assertEqual(self.matrix[(FIRST_ROW, "ANSYS Mechanical")], Status.MISSING.value)  # Required, no row

    def test_matrix_shows_not_applicable_and_optional_defaults_without_rows(self):
        casey = CASEY_ROW
        self.assertEqual(self.matrix[(casey, "Revit")], Status.NOT_APPLICABLE.value)
        self.assertEqual(self.matrix[(casey, "Microsoft Visio")], Status.NOT_INSTALLED.value)
        self.assertEqual(self.matrix[(casey, "Microsoft 365 Apps")], Status.MISSING.value)

    def test_matrix_uses_the_inventory_row_status_when_there_is_one(self):
        self.assertEqual(self.matrix[(FIRST_ROW, "SOLIDWORKS")], Status.OUTDATED.value)
        self.assertEqual(self.matrix[(CASEY_ROW, "AutoCAD")], Status.MISSING.value)  # override to Required

    def test_person_view_counts_missing_for_the_picked_person(self):
        self.assertGreaterEqual(self.person_view_missing, 1)


if __name__ == "__main__":
    unittest.main()
