"""Utilitas PDF: deteksi kemampuan, konversi DOCX→PDF (Word COM), gabung PDF.

Semua fungsi bebas dari dependensi Flet sehingga mudah dipakai ulang dari
skrip CLI, notebook, maupun aplikasi GUI.
"""
import os
import sys
from pathlib import Path


# Keep these handles alive for the lifetime of the process. Closing an
# os.add_dll_directory() handle removes its directory from the DLL search path.
_PYWIN32_DLL_DIRECTORIES = []


def configure_bundled_pywin32(search_roots=None):
    """Restore the path setup normally performed by ``pywin32.pth``.

    pywin32 installs ``pythoncom.py`` below ``win32`` rather than directly in
    site-packages. In a regular Python installation, ``pywin32.pth`` adds that
    directory (and ``win32/lib``) to ``sys.path`` and imports a bootstrap module
    that exposes pywin32's native DLLs. Flet copies the wheel payload into its
    embedded runtime but does not guarantee processing third-party ``.pth``
    files, so perform the equivalent setup explicitly before importing COM.

    ``search_roots`` is primarily for tests; production uses the embedded
    interpreter's existing import paths.
    """
    roots = search_roots if search_roots is not None else sys.path
    for root in roots:
        try:
            root_path = Path(root)
        except TypeError:
            continue

        # A normal site-packages root has win32/pythoncom.py. The fallback also
        # supports a Flet payload where sys.path points at a parent directory.
        candidates = [root_path / "win32"]
        if root_path.is_dir():
            try:
                candidates.extend(path.parent for path in root_path.rglob("pythoncom.py"))
            except OSError:
                continue

        for win32_dir in candidates:
            if not (win32_dir / "pythoncom.py").is_file():
                continue

            package_root = win32_dir.parent
            for path in (win32_dir, win32_dir / "lib", package_root / "pythonwin"):
                if path.is_dir() and str(path) not in sys.path:
                    sys.path.insert(0, str(path))

            system32_dir = package_root / "pywin32_system32"
            if system32_dir.is_dir() and hasattr(os, "add_dll_directory"):
                handle = os.add_dll_directory(str(system32_dir))
                _PYWIN32_DLL_DIRECTORIES.append(handle)
            return True
    return False

# Deteksi kemampuan konversi PDF (butuh MS Word terpasang + paket pywin32)
PDF_AVAILABLE = False
try:
    configure_bundled_pywin32()
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
