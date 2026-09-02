"""Estimasi durasi untuk fase konversi DOCX ke PDF di Langkah 5."""

import math
from decimal import Decimal, ROUND_CEILING


# document_id: (workload benchmark, durasi benchmark dalam detik)
CONVERSION_BENCHMARKS: dict[str, tuple[int, int]] = {
    "lampiran_spk_ppl": (6, 24),
    "lampiran_spk_pml": (2, 8),
    "bapp_ppl_t1": (299, 290),
    "bapp_ppl_t2": (299, 290),
    "bapp_pml_t1": (47, 47),
    "bapp_pml_t2": (47, 47),
    "bast_ppl": (299, 426),
    "bast_pml": (47, 61),
    "bukti_terima": (8, 10),
    "spp_ppl": (6, 33),
    "spp_t2_ppl": (6, 33),
    "spp_pml": (2, 12),
    "spp_t2_pml": (2, 12),
}


def conversion_workload(
    document_id: str,
    generated_file_count: int,
    recipient_count: int | None = None,
) -> int:
    """Kembalikan unit kerja konversi sesuai karakteristik dokumen.

    Umumnya satu unit kerja adalah satu DOCX. Bukti Terima merupakan satu master
    DOCX, sehingga unit kerjanya adalah jumlah baris penerima di dalam dokumen.
    """
    workload = recipient_count if document_id == "bukti_terima" else generated_file_count
    if workload is None:
        raise ValueError("Jumlah penerima wajib diberikan untuk Bukti Terima.")
    if workload < 0:
        raise ValueError("Workload konversi tidak boleh negatif.")
    return workload


def estimate_conversion_seconds(
    document_id: str,
    workload: int,
    safety_buffer: float = 0.20,
) -> int:
    """Hitung estimasi maksimum linear, termasuk buffer, dibulatkan ke atas."""
    if document_id not in CONVERSION_BENCHMARKS:
        raise ValueError(f"Benchmark konversi tidak tersedia untuk '{document_id}'.")
    if workload < 0:
        raise ValueError("Workload konversi tidak boleh negatif.")
    if safety_buffer < 0:
        raise ValueError("Safety buffer tidak boleh negatif.")

    prior_workload, prior_duration = CONVERSION_BENCHMARKS[document_id]
    estimate = (
        Decimal(prior_duration)
        / Decimal(prior_workload)
        * Decimal(workload)
        * (Decimal("1") + Decimal(str(safety_buffer)))
    )
    return int(estimate.to_integral_value(rounding=ROUND_CEILING))


def format_estimated_duration(seconds: float) -> str:
    """Format durasi estimasi bulat dalam bahasa Indonesia."""
    whole_seconds = max(0, math.ceil(seconds))
    minutes, remainder = divmod(whole_seconds, 60)
    if minutes == 0:
        return f"{remainder} detik"
    if remainder == 0:
        return f"{minutes} menit"
    return f"{minutes} menit {remainder} detik"


def conversion_estimate_messages(max_seconds: int, elapsed_seconds: float) -> tuple[str, str]:
    """Buat teks estimasi maksimum dan status sisa/kelebihan waktu."""
    maximum = (
        "Estimasi selesai maksimal dalam "
        f"{format_estimated_duration(max_seconds)}"
    )
    if elapsed_seconds > max_seconds:
        remaining = "Penyelesaian membutuhkan waktu tambahan…"
    else:
        remaining = (
            "Perkiraan sisa waktu: "
            f"{format_estimated_duration(max_seconds - elapsed_seconds)}"
        )
    return maximum, remaining