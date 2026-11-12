"""Small cell-value helper."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tool"))

from cell_values import number_if_plain


class NumberIfPlainTests(unittest.TestCase):
    def test_plain_numbers_become_numbers(self):
        self.assertEqual(number_if_plain("2026"), 2026)
        self.assertEqual(number_if_plain("3.12"), 3.12)

    def test_text_and_lossy_numbers_stay_text(self):
        self.assertEqual(number_if_plain("2025 R2"), "2025 R2")
        self.assertEqual(number_if_plain("3.10"), "3.10")      # would become 3.1
        self.assertEqual(number_if_plain("R2025b"), "R2025b")
        self.assertEqual(number_if_plain(""), "")


if __name__ == "__main__":
    unittest.main()
