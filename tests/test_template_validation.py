"""Tests for validation of edited DOCX templates in Step 2."""

import os
import tempfile
import unittest

from docx import Document

from src.document_generator import (
    extract_template_placeholders,
    validate_template_placeholders,
)


class TemplatePlaceholderValidationTests(unittest.TestCase):
    def _save_document(self, directory, name, configure_document):
        document = Document()
        configure_document(document)
        path = os.path.join(directory, name)
        document.save(path)
        return path

    def test_extracts_placeholders_from_split_runs_tables_and_header(self):
        with tempfile.TemporaryDirectory() as directory:
            def configure(document):
                paragraph = document.add_paragraph()
                paragraph.add_run("{{nama_")
                paragraph.add_run("lengkap}}")
                document.add_table(rows=1, cols=1).cell(0, 0).text = "{{nik}}"
                document.sections[0].header.paragraphs[0].text = "{{no_spk}}"

            path = self._save_document(directory, "template.docx", configure)

            self.assertEqual(
                extract_template_placeholders(path),
                {"nama_lengkap", "nik", "no_spk"},
            )

    def test_accepts_template_with_exact_expected_placeholders(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._save_document(
                directory,
                "edited.docx",
                lambda document: document.add_paragraph(
                    "{{nik}} untuk {{nama_lengkap}}"
                ),
            )

            result = validate_template_placeholders(
                path, {"nama_lengkap", "nik"}
            )

            self.assertTrue(result["is_valid"])
            self.assertEqual(result["missing"], [])
            self.assertEqual(result["unexpected"], [])

    def test_rejects_missing_and_unexpected_placeholders(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._save_document(
                directory,
                "invalid.docx",
                lambda document: document.add_paragraph("{{nik}} {{jabatan}}"),
            )

            result = validate_template_placeholders(
                path, {"nama_lengkap", "nik"}
            )

            self.assertFalse(result["is_valid"])
            self.assertEqual(result["missing"], ["nama_lengkap"])
            self.assertEqual(result["unexpected"], ["jabatan"])


if __name__ == "__main__":
    unittest.main()