"""Native unit tests for SIOMAY release metadata and update validation."""

import unittest

from src.release import is_newer_package_version
from src.updates import parse_update_manifest


class ReleaseVersionTests(unittest.TestCase):
    def test_only_newer_numeric_package_versions_are_accepted(self):
        self.assertFalse(is_newer_package_version("2026.1.3"))
        self.assertFalse(is_newer_package_version("2026.1.2.1.0.2"))
        self.assertTrue(is_newer_package_version("2026.1.4"))
        self.assertTrue(is_newer_package_version("2026.2.0.0"))
        self.assertFalse(is_newer_package_version("invalid"))


class UpdateManifestTests(unittest.TestCase):
    def test_newer_official_release_is_returned(self):
        update = parse_update_manifest({
            "display_version": "v2026.1.4",
            "package_version": "2026.1.4",
            "download_url": (
                "https://github.com/Mjulianfr001056/siomay-se26/releases/"
                "download/v2026.1.4/SIOMAY-v2026.1.4-windows.zip"
            ),
            "release_notes_url": (
                "https://github.com/Mjulianfr001056/siomay-se26/releases/tag/"
                "v2026.1.4"
            ),
        })

        self.assertIsNotNone(update)
        self.assertEqual(update.package_version, "2026.1.4")

    def test_current_version_returns_no_update(self):
        update = parse_update_manifest({
            "display_version": "v2026.1.3",
            "package_version": "2026.1.3",
            "download_url": "https://github.com/Mjulianfr001056/siomay-se26/releases",
            "release_notes_url": "https://github.com/Mjulianfr001056/siomay-se26/releases",
        })

        self.assertIsNone(update)

    def test_unofficial_release_url_is_rejected(self):
        with self.assertRaises(ValueError):
            parse_update_manifest({
                "display_version": "v2026.1.4",
                "package_version": "2026.1.4",
                "download_url": "https://example.invalid/SIOMAY-Setup.exe",
                "release_notes_url": "https://example.invalid/notes",
            })

    def test_deceptive_release_path_is_rejected(self):
        with self.assertRaises(ValueError):
            parse_update_manifest({
                "display_version": "v2026.1.4",
                "package_version": "2026.1.4",
                "download_url": (
                    "https://github.com/Mjulianfr001056/siomay-se26/"
                    "releases-unofficial/Setup.exe"
                ),
                "release_notes_url": (
                    "https://github.com/Mjulianfr001056/siomay-se26/"
                    "releases-unofficial/notes"
                ),
            })