"""Native unit tests for SIOMAY release metadata and update validation."""

import unittest
from unittest.mock import Mock, patch

import requests

from src.release import is_newer_package_version
from src.updates import check_for_update, fetch_release_changelog, parse_update_manifest


class ReleaseVersionTests(unittest.TestCase):
    def test_only_newer_numeric_package_versions_are_accepted(self):
        self.assertFalse(is_newer_package_version("2026.1.4"))
        self.assertFalse(is_newer_package_version("2026.1.2.1.0.2"))
        self.assertTrue(is_newer_package_version("2026.2.0.0"))
        self.assertFalse(is_newer_package_version("invalid"))


class UpdateManifestTests(unittest.TestCase):
    def test_newer_official_release_is_returned(self):
        with patch("src.updates.is_newer_package_version", return_value=True):
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
            "display_version": "v2026.1.4",
            "package_version": "2026.1.4",
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


class ReleaseChangelogTests(unittest.TestCase):
    RELEASE_URL = (
        "https://github.com/Mjulianfr001056/siomay-se26/releases/tag/v2026.2.0"
    )

    @patch("src.updates.requests.get")
    def test_fetches_markdown_body_from_matching_official_release(self, mock_get):
        response = Mock()
        response.json.return_value = {
            "tag_name": "v2026.2.0",
            "body": "## Yang baru\n\n- Perbaikan penting",
        }
        mock_get.return_value = response

        result = fetch_release_changelog(self.RELEASE_URL, timeout=3.0)

        self.assertEqual(result, "## Yang baru\n\n- Perbaikan penting")
        mock_get.assert_called_once_with(
            "https://api.github.com/repos/Mjulianfr001056/siomay-se26/"
            "releases/tags/v2026.2.0",
            headers={"Accept": "application/vnd.github+json"},
            timeout=3.0,
        )
        response.raise_for_status.assert_called_once_with()

    @patch("src.updates.requests.get")
    def test_empty_release_body_returns_no_changelog(self, mock_get):
        response = Mock()
        response.json.return_value = {"tag_name": "v2026.2.0", "body": "  "}
        mock_get.return_value = response

        self.assertIsNone(fetch_release_changelog(self.RELEASE_URL))

    def test_rejects_non_tag_release_page(self):
        with self.assertRaises(ValueError):
            fetch_release_changelog(
                "https://github.com/Mjulianfr001056/siomay-se26/releases"
            )

    @patch("src.updates.is_newer_package_version", return_value=True)
    @patch("src.updates.requests.get")
    def test_check_for_update_includes_release_changelog(self, mock_get, _mock_newer):
        manifest_response = Mock()
        manifest_response.content = (
            b'{"display_version":"v2026.2.0","package_version":"2026.2.0",'
            b'"download_url":"https://github.com/Mjulianfr001056/siomay-se26/'
            b'releases/tag/v2026.2.0","release_notes_url":"https://github.com/'
            b'Mjulianfr001056/siomay-se26/releases/tag/v2026.2.0"}'
        )
        notes_response = Mock()
        notes_response.json.return_value = {
            "tag_name": "v2026.2.0",
            "body": "- Lebih andal",
        }
        mock_get.side_effect = [manifest_response, notes_response]

        update = check_for_update(timeout=2.0)

        self.assertEqual(update.changelog, "- Lebih andal")
        self.assertEqual(mock_get.call_count, 2)

    @patch(
        "src.updates.fetch_release_changelog",
        side_effect=requests.ConnectionError("offline"),
    )
    @patch("src.updates.is_newer_package_version", return_value=True)
    @patch("src.updates.requests.get")
    def test_changelog_failure_does_not_hide_update(
        self, mock_get, _mock_newer, _mock_changelog
    ):
        manifest_response = Mock()
        manifest_response.content = (
            b'{"display_version":"v2026.2.0","package_version":"2026.2.0",'
            b'"download_url":"https://github.com/Mjulianfr001056/siomay-se26/'
            b'releases/tag/v2026.2.0","release_notes_url":"https://github.com/'
            b'Mjulianfr001056/siomay-se26/releases/tag/v2026.2.0"}'
        )
        mock_get.return_value = manifest_response

        update = check_for_update()

        self.assertIsNotNone(update)
        self.assertIsNone(update.changelog)