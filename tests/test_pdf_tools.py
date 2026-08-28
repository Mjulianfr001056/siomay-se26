"""Tests for LibreOffice DOCX-to-PDF conversion helpers."""
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from utils.pdf_tools import convert_docx_to_pdf, find_libreoffice


class LibreOfficeTests(unittest.TestCase):
    def test_find_prefers_bundled_libreoffice(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            app_root = Path(temporary_directory)
            soffice = app_root / "LibreOffice" / "program" / "soffice.com"
            soffice.parent.mkdir(parents=True)
            soffice.touch()
            with patch("utils.pdf_tools.sys.executable", str(app_root / "SIOMAY.exe")):
                self.assertEqual(find_libreoffice(), soffice)

    def test_convert_uses_isolated_profile_and_moves_output(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source.docx"
            destination = root / "result.pdf"
            soffice = root / "soffice.com"
            source.touch()
            soffice.touch()

            def fake_run(command, **kwargs):
                output_dir = Path(command[command.index("--outdir") + 1])
                (output_dir / "source.pdf").write_bytes(b"pdf")
                self.assertIn("--headless", command)
                self.assertTrue(any(arg.startswith("-env:UserInstallation=file:") for arg in command))
                self.assertEqual(kwargs["timeout"], 120)
                return subprocess.CompletedProcess(command, 0, "", "")

            with patch("utils.pdf_tools.find_libreoffice", return_value=soffice), \
                    patch("utils.pdf_tools.subprocess.run", side_effect=fake_run):
                convert_docx_to_pdf(str(source), str(destination))

            self.assertEqual(destination.read_bytes(), b"pdf")

    def test_convert_reports_missing_libreoffice(self):
        with patch("utils.pdf_tools.find_libreoffice", return_value=None):
            with self.assertRaisesRegex(RuntimeError, "LibreOffice tidak ditemukan"):
                convert_docx_to_pdf("missing.docx", "result.pdf")