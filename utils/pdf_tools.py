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
    """Kembalikan executable LibreOffice yang dibundel atau tersedia lokal.

    Urutan pencarian:
    1. Folder LibreOffice/ di samping sys.executable  (distribusi portable / frozen)
    2. Folder LibreOffice/ di root repo / direktori app.py (dev-run: python app.py)
    3. Jalur instalasi sistem Windows standar (Program Files)
    4. PATH (soffice.exe / soffice.com / soffice)
    """
    # 1) Samping executable Python / frozen .exe
    executable_dir = Path(os.path.abspath(sys.executable)).parent
    candidates = [
        executable_dir / "LibreOffice" / "program" / "soffice.exe",
        executable_dir / "LibreOffice" / "program" / "soffice.com",
    ]

    # 2) Root repo — saat dijalankan dengan `python app.py` dari direktori proyek
    #    __file__ adalah utils/pdf_tools.py → dua level naik = root repo.
    repo_root = Path(os.path.abspath(__file__)).parent.parent
    candidates += [
        repo_root / "LibreOffice" / "program" / "soffice.exe",
        repo_root / "LibreOffice" / "program" / "soffice.com",
    ]

    # 3) Program Files Windows (instalasi sistem)
    candidates += [
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "LibreOffice" / "program" / "soffice.exe",
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "LibreOffice" / "program" / "soffice.com",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
        / "LibreOffice" / "program" / "soffice.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
        / "LibreOffice" / "program" / "soffice.com",
    ]

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    # 4) PATH (utamakan soffice.exe agar tidak memunculkan jendela console)
    for command in ("soffice.exe", "soffice.com", "soffice"):
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


def _windows_subprocess_kwargs():
    """Kembalikan opsi subprocess agar LibreOffice tidak membuka jendela CMD."""
    if sys.platform != "win32":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "creationflags": (
            subprocess.CREATE_NO_WINDOW
            if hasattr(subprocess, "CREATE_NO_WINDOW")
            else 0x08000000
        ),
        "startupinfo": startupinfo,
    }


def convert_docx_files_to_pdf(sources, output_dir: str):
    """Konversi banyak DOCX dengan satu proses LibreOffice.

    Mengembalikan ``(converted, failed)``. ``converted`` adalah daftar jalur PDF
    dalam urutan input, sedangkan ``failed`` berisi jalur DOCX yang tidak
    menghasilkan PDF. Nama staging yang pendek mencegah command line Windows
    menjadi terlalu panjang ketika ratusan dokumen dikonversi sekaligus.
    """
    soffice = find_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice tidak ditemukan. Ekstrak seluruh isi rilis SIOMAY "
            "(termasuk folder LibreOffice) sebelum mengonversi PDF."
        )
    source_paths = [Path(source).resolve() for source in sources]
    if not source_paths:
        return [], []
    for source in source_paths:
        if not source.is_file():
            raise RuntimeError(f"Dokumen sumber tidak ditemukan: {source}")
    destination_dir = Path(output_dir).resolve()
    destination_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="siomay_lo_profile_") as profile_dir, \
            tempfile.TemporaryDirectory(prefix="siomay_lo_input_") as input_dir, \
            tempfile.TemporaryDirectory(prefix="siomay_lo_output_") as lo_output_dir:
        staged_sources = []
        for index, source in enumerate(source_paths, start=1):
            staged = Path(input_dir) / f"{index:06d}.docx"
            shutil.copy2(source, staged)
            staged_sources.append(staged)

        command = [
            str(soffice), "--headless", "--nologo", "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
            "--convert-to", "pdf:writer_pdf_Export", "--outdir", lo_output_dir,
            *[str(source) for source in staged_sources],
        ]

        try:
            result = subprocess.run(
                command, check=False, capture_output=True, text=True,
                encoding="utf-8", errors="replace",
                timeout=max(120, len(source_paths) * 10),
                **_windows_subprocess_kwargs(),
            )
        except OSError as ex:
            raise RuntimeError(f"Gagal menjalankan LibreOffice: {ex}") from ex
        except subprocess.TimeoutExpired as ex:
            raise RuntimeError(
                "Konversi PDF batch oleh LibreOffice melebihi batas waktu."
            ) from ex

        converted, failed = [], []
        for source, staged in zip(source_paths, staged_sources):
            staged_pdf = Path(lo_output_dir) / f"{staged.stem}.pdf"
            if not staged_pdf.is_file():
                failed.append(str(source))
                continue
            destination = destination_dir / f"{source.stem}.pdf"
            # Generator semestinya memberi nama unik; hindari overwrite diam-diam.
            if destination.exists():
                suffix = 2
                while (destination_dir / f"{source.stem}_{suffix}.pdf").exists():
                    suffix += 1
                destination = destination_dir / f"{source.stem}_{suffix}.pdf"
            shutil.move(str(staged_pdf), str(destination))
            converted.append(str(destination))

        if not converted:
            details = (result.stderr or result.stdout).strip()
            suffix = f" Detail LibreOffice: {details}" if details else ""
            raise RuntimeError(
                f"Tidak ada DOCX yang berhasil dikonversi dengan LibreOffice.{suffix}"
            )
        return converted, failed


def convert_docx_to_pdf(src: str, dst: str):
    """Konversi satu DOCX ke PDF melalui LibreOffice secara headless."""
    destination = Path(dst).resolve()
    converted, _ = convert_docx_files_to_pdf([src], str(destination.parent))
    generated = Path(converted[0])
    if generated != destination:
        if destination.exists():
            destination.unlink()
        shutil.move(str(generated), str(destination))


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
