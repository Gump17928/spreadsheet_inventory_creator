"""Consistency of the starter and demo data: the things that would silently produce a wrong workbook."""
import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tool"))

import demo_data
import demo_inventory
import starter_data
from layout import N_PEOPLE, N_RECORDS, N_ROLES, N_TITLES
from vocab import Applicability


class StarterDataTests(unittest.TestCase):
    def test_everything_fits_the_workbook_capacity(self):
        self.assertLessEqual(len(starter_data.TITLES), N_TITLES)
        self.assertLessEqual(len(starter_data.ROLES), N_ROLES)
        self.assertLessEqual(len(demo_data.PEOPLE), N_PEOPLE)

    def test_role_profiles_only_name_catalog_titles(self):
        names = set(starter_data.TITLE_NAMES)
        for role, profile in starter_data.ROLE_PROFILES.items():
            self.assertLessEqual(set(profile.required) | set(profile.optional), names, role)

    def test_no_title_is_both_required_and_optional(self):
        for role, profile in starter_data.ROLE_PROFILES.items():
            self.assertFalse(set(profile.required) & set(profile.optional), role)

    def test_title_names_are_unique(self):
        names = starter_data.TITLE_NAMES
        self.assertEqual(len(names), len(set(names)))

    def test_titles_use_known_groups_and_categories(self):
        for title in starter_data.TITLES:
            self.assertIn(title.group, starter_data.GROUPS, title.name)
            self.assertIn(title.category, starter_data.CATEGORIES, title.name)

    def test_role_default_covers_all_three_outcomes(self):
        self.assertEqual(starter_data.role_default("Business Admin", "Microsoft 365 Apps"), Applicability.REQUIRED)
        self.assertEqual(starter_data.role_default("Business Admin", "Microsoft Visio"), Applicability.OPTIONAL)
        self.assertEqual(starter_data.role_default("Business Admin", "AutoCAD"), Applicability.NOT_APPLICABLE)


class DemoDataTests(unittest.TestCase):
    def test_every_title_has_license_details(self):
        self.assertEqual(set(demo_data.LICENSE_DETAILS), set(starter_data.TITLE_NAMES))

    def test_every_demo_person_has_a_known_role_and_department(self):
        for person in demo_data.PEOPLE:
            self.assertIn(person.role, starter_data.ROLES, person.name)
            self.assertIn(person.department, starter_data.DEPARTMENTS, person.name)

    def test_demo_records_are_valid_and_unique(self):
        records = demo_inventory.make_records(date(2026, 1, 1))
        people = {person.name for person in demo_data.PEOPLE}
        keys = [(r.person, r.software) for r in records]
        self.assertEqual(len(keys), len(set(keys)), "duplicate person+software rows")
        self.assertLessEqual(len(records), N_RECORDS)
        for record in records:
            self.assertIn(record.person, people)
            self.assertIn(record.software, starter_data.TITLE_NAMES)

    def test_demo_records_are_reproducible(self):
        first = demo_inventory.make_records(date(2026, 1, 1))
        second = demo_inventory.make_records(date(2026, 1, 1))
        self.assertEqual(first, second)


if __name__ == "__main__":
    unittest.main()
