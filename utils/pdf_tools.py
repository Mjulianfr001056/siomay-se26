"""Utilitas PDF: deteksi kemampuan, konversi DOCX→PDF (Word COM), gabung PDF.

Semua fungsi bebas dari dependensi Flet sehingga mudah dipakai ulang dari
skrip CLI, notebook, maupun aplikasi GUI.
"""
import os

# Deteksi kemampuan konversi PDF (butuh MS Word terpasang + paket pywin32)
PDF_AVAILABLE = False
try:
    import pythoncom  # noqa: F401
    import win32com.client  # noqa: F401
    PDF_AVAILABLE = True
except Exception:
    pass

# Deteksi kemampuan penggabungan PDF (paket pypdf — murni Python)
MERGE_AVAILABLE = False
try:
    from pypdf import PdfWriter  # noqa: F401
    MERGE_AVAILABLE = True
except Exception:
    pass


def convert_docx_to_pdf(src: str, dst: str):
    """Konversi satu .docx → .pdf memakai Word COM (late binding).

    Sengaja TIDAK memakai docx2pdf/gencache.EnsureDispatch: cache makepy-nya
    (folder gen_py) yang korup/stale sering memicu error samar seperti
    'Word.Application.Documents'. Late binding lewat DispatchEx sepenuhnya
    melewati cache tersebut. Aman dipanggil dari worker thread karena COM
    di-initialize & di-uninitialize di dalam fungsi ini.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    word = None
    try:
        import win32com.client as win32
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
        try:
            # 17 = wdFormatPDF. SaveAs2 tersedia sejak Word 2010;
            # fallback ke SaveAs untuk versi lama.
            try:
                doc.SaveAs2(os.path.abspath(dst), FileFormat=17)
            except Exception:
                doc.SaveAs(os.path.abspath(dst), FileFormat=17)
        finally:
            doc.Close(False)
    except Exception as ex:
        raise RuntimeError(
            f"Gagal mengonversi '{os.path.basename(src)}' ke PDF: {ex}"
        ) from ex
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


def merge_pdfs(pdf_paths, dst: str):
    """Gabung beberapa PDF menjadi satu berkas (urutan sesuai daftar input)."""
    from pypdf import PdfWriter

    writer = PdfWriter()
    try:
        for p in pdf_paths:
            writer.append(p)
        writer.write(dst)
    finally:
        writer.close()
