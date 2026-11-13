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
from demo_inventory import PERSON_INDEX, SLOT_INDEX, InstallRecord
from layout import FIRST_ROW, Cat, Inv, Mat, Ppl, Sheet, inventory_row
from starter_data import TITLE_NAMES
from vocab import Active, Applicability, InstallState, LicenseAssign, Status

TODAY = date.today()
ALEX, CASEY, TAYLOR = "Alex Rivera", "Casey Morgan", "Taylor Brooks"   # Mechanical Eng, Business Admin x2
FRESH = TODAY - timedelta(days=10)
STALE = TODAY - timedelta(days=400)
INSTALLED, NOT_INSTALLED = InstallState.INSTALLED, InstallState.NOT_INSTALLED
NAMED, FLOATING = LicenseAssign.NAMED, LicenseAssign.FLOATING
NEW_PERSON, NEW_TITLE = "Pat Newhire", "NewTool"
NEW_PERSON_INDEX = len(PERSON_INDEX)           # first empty People row
NEW_TITLE_SLOT = len(SLOT_INDEX)               # first empty Catalog row
SPARE_SLOT = NEW_TITLE_SLOT + 1                # a Catalog slot nobody has filled


def record(software, state, version="", assignment="", verified=FRESH, person=ALEX, **extra):
    return InstallRecord(person, software, state, version, None, assignment, verified_on=verified, **extra)


# (typed record, expected status). Alex: AutoCAD Optional, Revit N/A, SOLIDWORKS and ANSYS Mechanical Required.
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
    (record("AutoCAD", NOT_INSTALLED, person=CASEY, applicability_override=Applicability.REQUIRED), Status.MISSING),
]
# Nothing typed at all: the status comes from the role default alone.
BLANK_CASES = [
    (ALEX, "ANSYS Mechanical", Status.MISSING), (ALEX, "Bluebeam Revu", Status.NOT_INSTALLED),
    (CASEY, "Microsoft 365 Apps", Status.MISSING), (CASEY, "Microsoft Visio", Status.NOT_INSTALLED),
    (CASEY, "Inventor", Status.NOT_APPLICABLE), (CASEY, "Revit", Status.NOT_APPLICABLE),
    (NEW_PERSON, "SOLIDWORKS", Status.MISSING), (NEW_PERSON, "Revit", Status.NOT_APPLICABLE),
    (ALEX, NEW_TITLE, Status.NOT_INSTALLED),      # a title added later: role grid cell is blank -> Optional
]


def person_slot_cell(person_index, slot, col):
    return f"{col}{inventory_row(person_index, slot)}"


def _position(person, software):
    """(person index, title slot); anyone or anything not in the demo data is the one added by the test."""
    return PERSON_INDEX.get(person, NEW_PERSON_INDEX), SLOT_INDEX.get(software, NEW_TITLE_SLOT)


def inventory_address(person, software, col):
    person_index, slot = _position(person, software)
    return person_slot_cell(person_index, slot, col)


def matrix_address(person, software):
    person_index, slot = _position(person, software)
    return f"{Mat.title_col(slot)}{FIRST_ROW + person_index}"


def build_scenarios(path):
    wb = build_workbook.create_structure()
    build_workbook.add_starter_content(wb)
    sheet_people.fill_demo(wb[Sheet.PEOPLE])
    sheet_inventory.write_records(wb[Sheet.INVENTORY], [case for case, _expected in CASES], PERSON_INDEX, SLOT_INDEX)
    people, catalog = wb[Sheet.PEOPLE], wb[Sheet.CATALOG]
    people[f"{Ppl.ACTIVE}{FIRST_ROW + PERSON_INDEX[TAYLOR]}"] = Active.NO
    new_row = FIRST_ROW + NEW_PERSON_INDEX
    people[f"{Ppl.NAME}{new_row}"], people[f"{Ppl.ROLE}{new_row}"] = NEW_PERSON, "Mechanical Engineer"
    people[f"{Ppl.DEPARTMENT}{new_row}"] = "Mechanical"
    new_slot = FIRST_ROW + NEW_TITLE_SLOT
    catalog[f"{Cat.NAME}{new_slot}"], catalog[f"{Cat.GROUP}{new_slot}"] = NEW_TITLE, "Other"
    wb.save(path)


@unittest.skipUnless(excel_com.excel_available(), "Excel is not available")
class ExcelStatusTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        path = Path(cls.tmp.name) / "scenarios.xlsx"
        build_scenarios(path)
        with excel_com.excel_session() as app, excel_com.open_workbook(app, path) as workbook:
            cls.errors = excel_com.error_cells(workbook)
            cls.typed = [cls._status(workbook, case.person, case.software) for case, _ in CASES]
            cls.blank = [cls._status(workbook, person, software) for person, software, _ in BLANK_CASES]
            cls.facts = cls._read_facts(workbook)

    @staticmethod
    def _status(workbook, person, software):
        return excel_com.read_cell(workbook, Sheet.INVENTORY, inventory_address(person, software, Inv.STATUS))

    @staticmethod
    def _read_facts(workbook):
        """Everything the non-status tests look at, read in one Excel session."""
        def get(sheet, address):
            return excel_com.read_cell(workbook, sheet, address)

        m365, solidworks = "Microsoft 365 Apps", "SOLIDWORKS"
        return {
            "people_count": get(Sheet.DASHBOARD, "A5"),
            "retired_matrix_name": get(Sheet.MATRIX, f"{Mat.PERSON}{FIRST_ROW + PERSON_INDEX[TAYLOR]}"),
            "retired_matrix_cell": get(Sheet.MATRIX, matrix_address(TAYLOR, m365)),
            "retired_inventory_status": get(Sheet.INVENTORY, inventory_address(TAYLOR, m365, Inv.STATUS)),
            "new_matrix_name": get(Sheet.MATRIX, f"{Mat.PERSON}{FIRST_ROW + NEW_PERSON_INDEX}"),
            "new_matrix_solidworks": get(Sheet.MATRIX, matrix_address(NEW_PERSON, solidworks)),
            "new_title_header": get(Sheet.MATRIX, f"{Mat.title_col(NEW_TITLE_SLOT)}5"),
            "spare_slot_person": get(Sheet.INVENTORY, person_slot_cell(0, SPARE_SLOT, Inv.PERSON)),
            "spare_slot_status": get(Sheet.INVENTORY, person_slot_cell(0, SPARE_SLOT, Inv.STATUS)),
            "alex_matrix_solidworks": get(Sheet.MATRIX, matrix_address(ALEX, solidworks)),
            "row_owner": get(Sheet.INVENTORY, person_slot_cell(PERSON_INDEX[CASEY], 1, Inv.PERSON)),
            "row_title": get(Sheet.INVENTORY, person_slot_cell(PERSON_INDEX[CASEY], 1, Inv.SOFTWARE)),
        }

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_no_formula_errors_anywhere(self):
        self.assertEqual(self.errors, {})

    def test_every_typed_status_is_calculated_correctly(self):
        for index, (case, expected) in enumerate(CASES):
            with self.subTest(software=case.software, person=case.person):
                self.assertEqual(self.typed[index], expected.value)

    def test_blank_rows_follow_the_role_default(self):
        for index, (person, software, expected) in enumerate(BLANK_CASES):
            with self.subTest(person=person, software=software):
                self.assertEqual(self.blank[index], expected.value)

    def test_each_row_belongs_to_its_own_person_and_title(self):
        self.assertEqual(self.facts["row_owner"], CASEY)
        self.assertEqual(self.facts["row_title"], TITLE_NAMES[1])

    def test_spare_title_slots_stay_blank(self):
        self.assertFalse(self.facts["spare_slot_person"])
        self.assertFalse(self.facts["spare_slot_status"])

    def test_matrix_reads_the_inventory_row(self):
        self.assertEqual(self.facts["alex_matrix_solidworks"], Status.OUTDATED.value)

    def test_a_new_person_and_a_new_title_appear_with_no_extra_work(self):
        self.assertEqual(self.facts["new_matrix_name"], NEW_PERSON)
        self.assertEqual(self.facts["new_matrix_solidworks"], Status.MISSING.value)
        self.assertEqual(self.facts["new_title_header"], NEW_TITLE)

    def test_a_retired_person_drops_out_of_the_matrix_and_counts(self):
        self.assertFalse(self.facts["retired_matrix_name"])
        self.assertFalse(self.facts["retired_matrix_cell"])
        self.assertFalse(self.facts["retired_inventory_status"])
        self.assertEqual(self.facts["people_count"], len(PERSON_INDEX))   # 14 - 1 retired + 1 new


if __name__ == "__main__":
    unittest.main()
