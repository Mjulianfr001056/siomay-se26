"""Regression tests for the shared PPL/PML BAST generator."""

import unittest
import inspect
import tempfile
from unittest.mock import patch

import pandas as pd

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

    def test_generator_orientation_defaults_to_portrait(self):
        parameters = inspect.signature(bast.iter_generate).parameters
        self.assertIn("image_orientation", parameters)
        self.assertEqual(
            parameters["image_orientation"].default,
            bast.IMAGE_ORIENTATION_PORTRAIT,
        )

    def test_generator_forwards_orientation_to_evidence_document_generation(self):
        dfs = {
            bast.SHEET_NAME: pd.DataFrame([
                {"nik": "123", "nama_mitra": "Mitra", "jabatan": "ppl"}
            ])
        }
        with tempfile.TemporaryDirectory() as output_dir:
            with patch.object(bast, "_generate_one_doc", return_value=[]) as generate:
                list(bast.iter_generate(
                    "ppl", dfs, "template.docx", output_dir,
                    image_layout="dedicated_pages",
                    image_orientation="automatic",
                ))

            args = generate.call_args.args
        self.assertEqual(args[-2], "dedicated_pages")
        self.assertEqual(args[-1], "automatic")


if __name__ == "__main__":
    unittest.main()