"""Unit tests for UI helper routines (duration formatters & components)."""
import unittest

import flet as ft

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


if __name__ == "__main__":
    unittest.main()
