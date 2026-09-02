"""Tests for Step 5 DOCX-to-PDF conversion estimates."""

import unittest

from utils import (
    CONVERSION_BENCHMARKS,
    conversion_estimate_messages,
    conversion_workload,
    estimate_conversion_seconds,
    format_estimated_duration,
)


class ConversionEstimationTests(unittest.TestCase):
    def test_every_document_id_has_the_expected_benchmark(self):
        self.assertEqual(CONVERSION_BENCHMARKS, {
            "lampiran_spk_ppl": (6, 24),
            "lampiran_spk_pml": (2, 8),
            "bapp_ppl_t1": (299, 290),
            "bapp_ppl_t2": (6, 26),
            "bapp_pml_t1": (47, 47),
            "bapp_pml_t2": (47, 47),
            "bast_ppl": (299, 426),
            "bast_pml": (47, 61),
            "bukti_terima": (8, 10),
            "spp_ppl": (6, 33),
            "spp_t2_ppl": (6, 33),
            "spp_pml": (2, 12),
            "spp_t2_pml": (2, 12),
        })

    def test_each_benchmark_receives_twenty_percent_buffer_and_rounds_up(self):
        expected = {
            "lampiran_spk_ppl": 29,
            "lampiran_spk_pml": 10,
            "bapp_ppl_t1": 348,
            "bapp_ppl_t2": 348,
            "bapp_pml_t1": 57,
            "bapp_pml_t2": 57,
            "bast_ppl": 512,
            "bast_pml": 74,
            "bukti_terima": 12,
            "spp_ppl": 40,
            "spp_t2_ppl": 40,
            "spp_pml": 15,
            "spp_t2_pml": 15,
        }
        for document_id, estimate in expected.items():
            with self.subTest(document_id=document_id):
                prior_workload, _ = CONVERSION_BENCHMARKS[document_id]
                self.assertEqual(
                    estimate_conversion_seconds(document_id, prior_workload),
                    estimate,
                )

    def test_rounding_is_always_up_to_a_whole_second(self):
        self.assertEqual(estimate_conversion_seconds("lampiran_spk_ppl", 1), 5)
        self.assertEqual(
            estimate_conversion_seconds("lampiran_spk_ppl", 1, safety_buffer=0),
            4,
        )

    def test_workload_uses_recipient_rows_only_for_bukti_terima(self):
        self.assertEqual(conversion_workload("bukti_terima", 1, 27), 27)
        self.assertEqual(conversion_workload("bast_ppl", 12, 27), 12)
        with self.assertRaises(ValueError):
            conversion_workload("bukti_terima", 1)

    def test_estimated_duration_formatting(self):
        self.assertEqual(format_estimated_duration(30), "30 detik")
        self.assertEqual(format_estimated_duration(60), "1 menit")
        self.assertEqual(format_estimated_duration(135), "2 menit 15 detik")
        self.assertEqual(format_estimated_duration(18.1), "19 detik")

    def test_messages_show_static_maximum_and_count_down(self):
        maximum, remaining = conversion_estimate_messages(30, 12)
        self.assertEqual(maximum, "Estimasi selesai maksimal dalam 30 detik")
        self.assertEqual(remaining, "Perkiraan sisa waktu: 18 detik")

    def test_overdue_message_replaces_remaining_time(self):
        maximum, remaining = conversion_estimate_messages(30, 30.1)
        self.assertEqual(maximum, "Estimasi selesai maksimal dalam 30 detik")
        self.assertEqual(remaining, "Penyelesaian membutuhkan waktu tambahan…")


if __name__ == "__main__":
    unittest.main()