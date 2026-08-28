"""Tests for portable pywin32 setup used by Word PDF conversion."""
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from utils.pdf_tools import configure_bundled_pywin32


class BundledPywin32Tests(unittest.TestCase):
    def test_configure_adds_pywin32_module_and_dll_paths(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            package_root = Path(temporary_directory) / "site-packages"
            win32 = package_root / "win32"
            (win32 / "lib").mkdir(parents=True)
            (package_root / "pythonwin").mkdir()
            (package_root / "pywin32_system32").mkdir()
            (win32 / "pythoncom.py").touch()

            original_sys_path = sys.path[:]
            try:
                with patch.object(os, "add_dll_directory", create=True) as add_dll_directory:
                    self.assertTrue(configure_bundled_pywin32([str(package_root)]))

                self.assertIn(str(win32), sys.path)
                self.assertIn(str(win32 / "lib"), sys.path)
                self.assertIn(str(package_root / "pythonwin"), sys.path)
                add_dll_directory.assert_called_once_with(
                    str(package_root / "pywin32_system32")
                )
            finally:
                sys.path[:] = original_sys_path

    def test_configure_returns_false_when_pythoncom_is_not_bundled(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            self.assertFalse(configure_bundled_pywin32([temporary_directory]))