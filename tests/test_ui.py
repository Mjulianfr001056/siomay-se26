"""Unit tests for UI and files helper routines."""
import os
import unittest
from unittest.mock import MagicMock, patch

import flet as ft

from utils.files import open_external_url
from utils.ui import duration_info_box, format_duration, format_timer_clock


class DurationHelperTests(unittest.TestCase):
    def test_format_duration_sub_ten_seconds(self):
        self.assertEqual(format_duration(4.2), "4,2 detik")
        self.assertEqual(format_duration(0.0), "0,0 detik")
        self.assertEqual(format_duration(9.9), "9,9 detik")

    def test_format_duration_seconds(self):
        self.assertEqual(format_duration(12.4), "12 detik")
        self.assertEqual(format_duration(45.0), "45 detik")
        self.assertEqual(format_duration(59.4), "59 detik")

    def test_format_duration_minutes_and_seconds(self):
        self.assertEqual(format_duration(60.0), "1 menit")
        self.assertEqual(format_duration(65.0), "1 menit 5 detik")
        self.assertEqual(format_duration(120.0), "2 menit")
        self.assertEqual(format_duration(128.0), "2 menit 8 detik")

    def test_format_timer_clock(self):
        self.assertEqual(format_timer_clock(0), "00:00")
        self.assertEqual(format_timer_clock(5), "00:05")
        self.assertEqual(format_timer_clock(65), "01:05")
        self.assertEqual(format_timer_clock(605), "10:05")

    def test_duration_info_box(self):
        box = duration_info_box(
            title="Waktu Pembuatan Dokumen Selesai",
            items=[
                ("Dokumen", "Lampiran SPK (10 berkas)"),
                ("Waktu pembuatan (generate)", "12 detik"),
            ],
            icon=ft.Icons.TIMER_ROUNDED,
        )
        self.assertEqual(box.bgcolor, ft.Colors.PURPLE_50)
        self.assertIsInstance(box.content, ft.Column)
        # Column has 1 title row + 2 items rows
        self.assertEqual(len(box.content.controls), 3)


class ExternalUrlHelperTests(unittest.TestCase):
    @patch("webbrowser.open", return_value=True)
    def test_open_external_url_webbrowser_success(self, mock_webbrowser):
        open_external_url("https://github.com/example")
        mock_webbrowser.assert_called_once_with("https://github.com/example")

    @patch("webbrowser.open", side_effect=Exception("Failed"))
    @patch("os.startfile", create=True)
    def test_open_external_url_page_fallback(self, mock_startfile, mock_webbrowser):
        mock_page = MagicMock()
        open_external_url("https://github.com/example", page=mock_page)
        mock_page.launch_url.assert_called_once_with("https://github.com/example")

    @patch("webbrowser.open", side_effect=Exception("Failed"))
    def test_open_external_url_os_startfile_fallback(self, mock_webbrowser):
        mock_page = MagicMock()
        mock_page.launch_url.side_effect = Exception("Failed")
        with patch("os.startfile", create=True) as mock_startfile:
            open_external_url("https://github.com/example", page=mock_page)
            mock_startfile.assert_called_once_with("https://github.com/example")

    def test_open_external_url_empty(self):
        # Should do nothing safely without exception
        open_external_url("")
        open_external_url(None)


if __name__ == "__main__":
    unittest.main()
