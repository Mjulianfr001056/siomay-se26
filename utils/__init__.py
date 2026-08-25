"""utils — pustaka lokal berisi fungsi-fungsi utilitas yang dapat dipakai ulang.

Submodul:
- pdf_tools : deteksi kemampuan, konversi DOCX→PDF (Word COM), gabung PDF (pypdf)
- files     : urutan alami nama berkas, ekstensi, arsip ZIP, buka Explorer,
              opsi dialog simpan
- ui        : helper Flet (snackbar, log aktivitas, stat_box, tutup jendela)
"""
from utils.files import (
    ensure_extension,
    file_order_key,
    open_in_explorer,
    save_dialog_options,
    zip_files,
)
from utils.pdf_tools import (
    MERGE_AVAILABLE,
    PDF_AVAILABLE,
    convert_docx_to_pdf,
    merge_pdfs,
)
from utils.ui import (
    LOG_STYLES,
    close_window,
    make_activity_log,
    make_snackbar,
    stat_box,
)

__all__ = [
    "MERGE_AVAILABLE",
    "PDF_AVAILABLE",
    "LOG_STYLES",
    "close_window",
    "convert_docx_to_pdf",
    "ensure_extension",
    "file_order_key",
    "make_activity_log",
    "make_snackbar",
    "merge_pdfs",
    "open_in_explorer",
    "save_dialog_options",
    "stat_box",
    "zip_files",
]
