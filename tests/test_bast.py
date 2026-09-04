"""Regression tests for the shared PPL/PML BAST generator."""

import unittest
import inspect

from src import bast


class BastSequenceNumberTests(unittest.TestCase):
    def test_sequence_number_preserves_three_digit_form(self):
        self.assertEqual(bast._format_no_urut_bast("001"), "001")

    def test_numeric_excel_value_is_restored_to_three_digits(self):
        self.assertEqual(bast._format_no_urut_bast(1), "001")
        self.assertEqual(bast._format_no_urut_bast("21"), "021")

    def test_non_numeric_sequence_number_is_unchanged(self):
        self.assertEqual(bast._format_no_urut_bast("BAST-1"), "BAST-1")
        self.assertEqual(bast._format_no_urut_bast(""), "")


class BastEvidenceLayoutTests(unittest.TestCase):
    def test_generator_accepts_image_layout_option(self):
        parameters = inspect.signature(bast.iter_generate).parameters
        self.assertIn("image_layout", parameters)
        self.assertEqual(
            parameters["image_layout"].default,
            bast.IMAGE_LAYOUT_GRID,
        )


if __name__ == "__main__":
    unittest.main()