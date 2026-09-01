"""Tests for LibreOffice DOCX-to-PDF conversion helpers."""
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from utils.pdf_tools import (
    convert_docx_files_to_pdf,
    convert_docx_to_pdf,
    find_libreoffice,
)


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
                (output_dir / "000001.pdf").write_bytes(b"pdf")
                self.assertIn("--headless", command)
                self.assertTrue(any(arg.startswith("-env:UserInstallation=file:") for arg in command))
                self.assertEqual(kwargs["timeout"], 120)
                return subprocess.CompletedProcess(command, 0, "", "")

            with patch("utils.pdf_tools.find_libreoffice", return_value=soffice), \
                    patch("utils.pdf_tools.subprocess.run", side_effect=fake_run):
                convert_docx_to_pdf(str(source), str(destination))

            self.assertEqual(destination.read_bytes(), b"pdf")

    def test_batch_conversion_starts_libreoffice_once_and_reports_failures(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            sources = [root / "alpha.docx", root / "beta.docx", root / "gamma.docx"]
            for source in sources:
                source.touch()
            output_dir = root / "pdf"
            soffice = root / "soffice.com"
            soffice.touch()

            def fake_run(command, **kwargs):
                lo_output = Path(command[command.index("--outdir") + 1])
                # Simulasikan file kedua gagal; file pertama dan ketiga berhasil.
                (lo_output / "000001.pdf").write_bytes(b"alpha-pdf")
                (lo_output / "000003.pdf").write_bytes(b"gamma-pdf")
                staged_docx = [Path(arg) for arg in command if arg.endswith(".docx")]
                self.assertEqual(len(staged_docx), 3)
                self.assertEqual(kwargs["timeout"], 120)
                return subprocess.CompletedProcess(command, 0, "", "")

            with patch("utils.pdf_tools.find_libreoffice", return_value=soffice), \
                    patch("utils.pdf_tools.subprocess.run", side_effect=fake_run) as run:
                converted, failed = convert_docx_files_to_pdf(sources, output_dir)

            run.assert_called_once()
            self.assertEqual(
                [Path(path).name for path in converted],
                ["alpha.pdf", "gamma.pdf"],
            )
            self.assertEqual(failed, [str(sources[1].resolve())])
            self.assertEqual((output_dir / "alpha.pdf").read_bytes(), b"alpha-pdf")
            self.assertEqual((output_dir / "gamma.pdf").read_bytes(), b"gamma-pdf")

    def test_convert_reports_missing_libreoffice(self):
        with patch("utils.pdf_tools.find_libreoffice", return_value=None):
            with self.assertRaisesRegex(RuntimeError, "LibreOffice tidak ditemukan"):
                convert_docx_to_pdf("missing.docx", "result.pdf")