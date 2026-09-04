"""utils — pustaka lokal berisi fungsi-fungsi utilitas yang dapat dipakai ulang.

Submodul:
 - conversion_estimation : benchmark dan estimasi waktu konversi DOCX→PDF
 - pdf_tools : deteksi kemampuan, konversi DOCX→PDF (LibreOffice), gabung PDF (pypdf)
- files     : urutan alami nama berkas, ekstensi, arsip ZIP, buka Explorer,
              opsi dialog simpan
- ui        : helper Flet (snackbar, log aktivitas, stat_box, tutup jendela)
"""
from utils.files import (
    ensure_extension,
    file_order_key,
    open_external_url,
    open_in_explorer,
    save_dialog_options,
    zip_files,
)
from utils.conversion_estimation import (
    CONVERSION_BENCHMARKS,
    conversion_estimate_messages,
    conversion_workload,
    estimate_conversion_seconds,
    format_estimated_duration,
)
from utils.debounce import DebounceGate
from utils.feedback import (
    FEEDBACK_URL,
    dismiss_feedback_prompt,
    record_launch_and_should_prompt,
)
from utils.pdf_tools import (
    MERGE_AVAILABLE,
    PDF_AVAILABLE,
    convert_docx_files_to_pdf,
    convert_docx_to_pdf,
    merge_pdfs,
)
from utils.ui import (
    LOG_STYLES,
    close_window,
    duration_info_box,
    format_duration,
    format_timer_clock,
    make_activity_log,
    make_snackbar,
    stat_box,
)

__all__ = [
    "CONVERSION_BENCHMARKS",
    "DebounceGate",
    "FEEDBACK_URL",
    "MERGE_AVAILABLE",
    "PDF_AVAILABLE",
    "LOG_STYLES",
    "close_window",
    "conversion_estimate_messages",
    "conversion_workload",
    "convert_docx_files_to_pdf",
    "convert_docx_to_pdf",
    "duration_info_box",
    "dismiss_feedback_prompt",
    "ensure_extension",
    "estimate_conversion_seconds",
    "file_order_key",
    "format_duration",
    "format_estimated_duration",
    "format_timer_clock",
    "make_activity_log",
    "make_snackbar",
    "merge_pdfs",
    "open_external_url",
    "open_in_explorer",
    "record_launch_and_should_prompt",
    "save_dialog_options",
    "stat_box",
    "zip_files",
]
