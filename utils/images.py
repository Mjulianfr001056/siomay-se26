"""Pengunduhan dan normalisasi gambar untuk penyisipan ke dokumen Word."""

from __future__ import annotations

import io
import time

import requests

try:
    from PIL import Image, ImageFile, ImageOps, UnidentifiedImageError

    ImageFile.LOAD_TRUNCATED_IMAGES = True
    HAS_PIL = True
except ImportError:  # pragma: no cover - dependency wajib pada paket aplikasi
    HAS_PIL = False

try:
    from pillow_heif import register_heif_opener

    register_heif_opener()
    HAS_HEIF = True
except (ImportError, OSError):  # OSError mencakup kegagalan memuat codec native
    HAS_HEIF = False


_DRIVE_DOWNLOAD_URL = "https://drive.google.com/uc"
_DRIVE_DIRECT_URL = "https://drive.usercontent.google.com/download"


def _looks_like_html(raw_bytes: bytes, content_type: str = "") -> bool:
    """Deteksi halaman HTML yang kadang dikembalikan Google Drive."""
    if "text/html" in (content_type or "").lower():
        return True
    prefix = raw_bytes[:512].lstrip().lower()
    return prefix.startswith((b"<!doctype html", b"<html"))


def image_bytes_to_png(raw_bytes: bytes, content_type: str = ""):
    """Decode gambar umum/HEIC lalu kembalikan ``(BytesIO PNG, PIL.Image)``.

    HEIC/HEIF dikenali oleh Pillow setelah ``pillow-heif`` mendaftarkan
    pluginnya. Gambar ditranspose mengikuti EXIF orientation dan dikonversi ke
    mode yang aman bagi PNG/python-docx sebelum disisipkan.
    """
    if not HAS_PIL:
        raise RuntimeError("Pillow tidak terinstal.")
    if not raw_bytes:
        raise RuntimeError("File gambar kosong.")
    if _looks_like_html(raw_bytes, content_type):
        raise RuntimeError(
            "Google Drive mengembalikan halaman HTML, bukan file gambar. "
            "Pastikan akses file disetel untuk siapa saja yang memiliki tautan."
        )

    try:
        with Image.open(io.BytesIO(raw_bytes)) as source:
            source.load()
            image = ImageOps.exif_transpose(source)
            image.load()
    except UnidentifiedImageError as exc:
        hint = (
            " Decoder HEIC/HEIF tidak tersedia."
            if not HAS_HEIF
            else " File mungkin rusak atau bukan format gambar yang didukung."
        )
        raise RuntimeError(f"Format gambar tidak dapat dikenali.{hint}") from exc

    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")

    png_file = io.BytesIO()
    image.save(png_file, format="PNG")
    png_file.seek(0)
    return png_file, image


def download_drive_image(
    file_id: str,
    *,
    max_retries: int = 3,
    retry_delay: float = 2,
    timeout: float = 30,
):
    """Unduh gambar publik Google Drive dan normalisasi hasilnya menjadi PNG."""
    if not HAS_PIL:
        raise RuntimeError("Pillow tidak terinstal.")

    session = requests.Session()
    try:
        for attempt in range(1, max_retries + 1):
            try:
                response = session.get(
                    _DRIVE_DOWNLOAD_URL,
                    params={"export": "download", "id": file_id},
                    stream=True,
                    timeout=timeout,
                )
                response.raise_for_status()

                if _looks_like_html(
                    response.content, response.headers.get("Content-Type", "")
                ):
                    token = next(
                        (
                            value
                            for key, value in response.cookies.items()
                            if key.startswith("download_warning")
                        ),
                        None,
                    )
                    if token:
                        response = session.get(
                            _DRIVE_DOWNLOAD_URL,
                            params={
                                "export": "download",
                                "id": file_id,
                                "confirm": token,
                            },
                            stream=True,
                            timeout=timeout,
                        )
                    else:
                        response = session.get(
                            _DRIVE_DIRECT_URL,
                            params={
                                "id": file_id,
                                "export": "download",
                                "confirm": "t",
                            },
                            stream=True,
                            timeout=timeout,
                        )
                    response.raise_for_status()

                return image_bytes_to_png(
                    response.content, response.headers.get("Content-Type", "")
                )
            except requests.RequestException as exc:
                if attempt == max_retries:
                    raise RuntimeError(
                        f"Gagal mengunduh file {file_id} setelah "
                        f"{max_retries} percobaan: {exc}"
                    ) from exc
                time.sleep(retry_delay)
    finally:
        session.close()