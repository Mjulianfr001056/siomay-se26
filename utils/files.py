"""Utilitas berkas — pengurutan alami, ekstensi, arsip ZIP, integrasi OS.

Semua fungsi murni (tanpa UI) sehingga dapat dipakai dari aplikasi GUI,
skrip CLI, maupun notebook.
"""
import os
import re
import webbrowser
import zipfile


def file_order_key(path: str):
    """Kunci urutan berdasar nomor urut di awal nama berkas.

    Generator menamai berkas '{no_urut or idx}_{nama}.docx' (Lampiran SPK)
    atau '{prefix}_{NNN}_{nama}.docx' (jalur generik), sehingga prefix
    numerik pada nama berkas = no_urut_spk / urutan baris data.
    Angka dibandingkan secara numerik (2 < 10, "002" == 2); berkas tanpa
    prefix angka diletakkan terakhir secara alfabetis.
    """
    name = os.path.basename(path)
    m = re.match(r"^(\d+)", name)
    if m:
        return (0, int(m.group(1)), name)
    return (1, 0, name)


def ensure_extension(path: str, ext: str) -> str:
    """Kembalikan `path` dengan ekstensi `ext` (tanpa menimpa yang sudah benar).

    Contoh: ensure_extension("hasil", "pdf") -> "hasil.pdf"
            ensure_extension("a.PDF", "pdf") -> "a.PDF"  (sudah sesuai)
    """
    ext = ext.lstrip(".")
    return path if path.lower().endswith("." + ext.lower()) else f"{path}.{ext}"


def zip_files(paths, dst: str):
    """Kumpulkan berkas ke arsip ZIP baru (deflated).

    Nama di dalam arsip = nama berkasnya saja (tanpa folder).
    Mengembalikan daftar nama arsip internal sesuai urutan penulisan,
    berguna untuk logging oleh pemanggil.
    """
    written = []
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            arcname = os.path.basename(p)
            zf.write(p, arcname=arcname)
            written.append(arcname)
    return written


def open_in_explorer(path: str):
    """Buka lokasi `path` (file atau folder) di Windows Explorer."""
    target = path if os.path.isdir(path) else os.path.dirname(path)
    if target and os.path.exists(target):
        os.startfile(target)  # Windows


def open_external_url(url: str, page=None):
    """Buka URL di peramban (browser) web bawaan pengguna dengan multi-fallback."""
    if not url:
        return
    # 1. Coba browser default Python
    try:
        if webbrowser.open(url):
            return
    except Exception:
        pass
    # 2. Coba page.launch_url jika tersedia
    if page is not None:
        try:
            page.launch_url(url)
            return
        except Exception:
            pass
    # 3. Fallback Windows OS startfile
    try:
        os.startfile(url)
    except Exception:
        pass


def save_dialog_options(fmt: str, prefix: str, stamp: str):
    """Saran nama berkas & daftar ekstensi untuk dialog simpan hasil.

    `fmt` salah satu dari: "zip" | "merged".
    Mengembalikan tuple (nama_berkas_saran, [ekstensi]).
    """
    if fmt == "merged":
        return f"{prefix}_gabung_{stamp}.pdf", ["pdf"]
    return f"{prefix}_{stamp}.zip", ["zip"]
