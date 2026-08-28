"""Utilitas PDF: konversi DOCX→PDF (LibreOffice), gabung PDF.

Semua fungsi bebas dari dependensi Flet sehingga mudah dipakai ulang dari
skrip CLI, notebook, maupun aplikasi GUI.
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_libreoffice() -> Path | None:
    """Kembalikan executable LibreOffice yang dibundel atau tersedia lokal."""
    executable_dir = Path(sys.executable).resolve().parent
    candidates = [
        executable_dir / "LibreOffice" / "program" / "soffice.com",
        executable_dir / "LibreOffice" / "program" / "soffice.exe",
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "LibreOffice" / "program" / "soffice.com",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
        / "LibreOffice" / "program" / "soffice.com",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    for command in ("soffice.com", "soffice.exe", "soffice"):
        location = shutil.which(command)
        if location:
            return Path(location)
    return None


# LibreOffice is a native application, bundled beside the Windows executable.
PDF_AVAILABLE = find_libreoffice() is not None

# Deteksi kemampuan penggabungan PDF (paket pypdf — murni Python)
MERGE_AVAILABLE = False
try:
    from pypdf import PdfWriter  # noqa: F401
    MERGE_AVAILABLE = True
except Exception:
    pass


def convert_docx_to_pdf(src: str, dst: str):
    """Konversi satu DOCX ke PDF melalui LibreOffice secara headless."""
    soffice = find_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice tidak ditemukan. Ekstrak seluruh isi rilis SIOMAY "
            "(termasuk folder LibreOffice) sebelum mengonversi PDF."
        )
    source = Path(src).resolve()
    destination = Path(dst).resolve()
    if not source.is_file():
        raise RuntimeError(f"Dokumen sumber tidak ditemukan: {source}")
    destination.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="siomay_lo_profile_") as profile_dir, \
            tempfile.TemporaryDirectory(prefix="siomay_lo_output_") as output_dir:
        command = [
            str(soffice), "--headless", "--nologo", "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
            "--convert-to", "pdf:writer_pdf_Export", "--outdir", output_dir,
            str(source),
        ]
        try:
            result = subprocess.run(
                command, check=False, capture_output=True, text=True,
                encoding="utf-8", errors="replace", timeout=120,
            )
        except OSError as ex:
            raise RuntimeError(f"Gagal menjalankan LibreOffice: {ex}") from ex
        except subprocess.TimeoutExpired as ex:
            raise RuntimeError("Konversi PDF oleh LibreOffice melebihi 120 detik.") from ex

        converted = Path(output_dir) / f"{source.stem}.pdf"
        if result.returncode != 0 or not converted.is_file():
            details = (result.stderr or result.stdout).strip()
            suffix = f" Detail LibreOffice: {details}" if details else ""
            raise RuntimeError(
                f"Gagal mengonversi '{source.name}' ke PDF dengan LibreOffice.{suffix}"
            )
        shutil.move(str(converted), str(destination))


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
