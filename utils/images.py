"""Pengunduhan dan normalisasi gambar/PDF untuk penyisipan ke dokumen Word."""

from __future__ import annotations

import io
import re
import time
from urllib.parse import urlparse

import requests

try:
    import pymupdf as fitz

    HAS_PDF_RENDERER = True
except ImportError:  # pragma: no cover - dependency wajib pada paket aplikasi
    HAS_PDF_RENDERER = False

try:
    from PIL import Image, ImageFile, ImageOps, UnidentifiedImageError

    ImageFile.LOAD_TRUNCATED_IMAGES = True
    HAS_PIL = True
except ImportError:  # pragma: no cover - dependency wajib pada paket aplikasi
    HAS_PIL = False

HEIF_IMPORT_ERROR = None

try:
    from pillow_heif import open_heif, register_heif_opener

    register_heif_opener()
    HAS_HEIF = True
except (ImportError, OSError) as exc:  # OSError: codec native gagal dimuat
    HAS_HEIF = False
    HEIF_IMPORT_ERROR = exc


_DRIVE_DOWNLOAD_URL = "https://drive.google.com/uc"
_DRIVE_DIRECT_URL = "https://drive.usercontent.google.com/download"
MAX_WEB_IMAGE_BYTES = 25 * 1024 * 1024


def extract_drive_file_id(link: str):
    """Extract a file ID from the common public Google Drive URL formats."""
    value = str(link or "").strip()
    patterns = (
        r"/file/d/([^/?#]+)",
        r"[?&]id=([^&#]+)",
        r"/d/([^/?#]+)",
    )
    for pattern in patterns:
        match = re.search(pattern, value)
        if match:
            return match.group(1)
    return None


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
        # Fallback langsung membuat dukungan HEIC tidak hanya bergantung pada
        # registrasi plugin Pillow dan memastikan freezer mendeteksi API native.
        if HAS_HEIF:
            try:
                heif_file = open_heif(raw_bytes, convert_hdr_to_8bit=True)
                image = heif_file.to_pillow()
                image = ImageOps.exif_transpose(image)
                image.load()
            except Exception as heif_exc:
                raise RuntimeError(
                    "Format gambar tidak dapat dikenali. File mungkin rusak "
                    "atau bukan format gambar yang didukung."
                ) from heif_exc
        else:
            detail = f" ({HEIF_IMPORT_ERROR})" if HEIF_IMPORT_ERROR else ""
            raise RuntimeError(
                "Format gambar tidak dapat dikenali. Decoder HEIC/HEIF tidak "
                "tersedia. Instal ulang aplikasi/dependensi pillow-heif"
                f"{detail}."
            ) from exc

    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")

    png_file = io.BytesIO()
    image.save(png_file, format="PNG")
    png_file.seek(0)
    return png_file, image


def pdf_bytes_to_png_pages(raw_bytes: bytes):
    """Render seluruh halaman PDF menjadi daftar ``(BytesIO PNG, ukuran)``."""
    if not HAS_PDF_RENDERER:
        raise RuntimeError("PyMuPDF tidak terinstal - file PDF tidak dapat diproses.")
    if not raw_bytes:
        raise RuntimeError("File PDF kosong.")

    pages = []
    try:
        with fitz.open(stream=raw_bytes, filetype="pdf") as pdf:
            if pdf.page_count == 0:
                raise RuntimeError("File PDF tidak memiliki halaman.")
            for page in pdf:
                # 144 dpi cukup tajam untuk DOCX tanpa membuat hasil terlalu besar.
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                png_file = io.BytesIO(pixmap.tobytes("png"))
                png_file.seek(0)
                pages.append((png_file, (pixmap.width, pixmap.height)))
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"File PDF tidak dapat dibaca: {exc}") from exc
    return pages


def _download_drive_bytes(file_id: str, max_retries: int, retry_delay: float,
                          timeout: float):
    """Unduh byte file publik Google Drive beserta Content-Type respons."""
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
                        (value for key, value in response.cookies.items()
                         if key.startswith("download_warning")),
                        None,
                    )
                    if token:
                        params = {"export": "download", "id": file_id,
                                  "confirm": token}
                        url = _DRIVE_DOWNLOAD_URL
                    else:
                        params = {"id": file_id, "export": "download",
                                  "confirm": "t"}
                        url = _DRIVE_DIRECT_URL
                    response = session.get(
                        url, params=params, stream=True, timeout=timeout
                    )
                    response.raise_for_status()

                content_type = response.headers.get("Content-Type", "")
                if _looks_like_html(response.content, content_type):
                    raise RuntimeError(
                        "Google Drive mengembalikan halaman HTML, bukan file. "
                        "Pastikan akses file disetel untuk siapa saja yang memiliki tautan."
                    )
                return response.content, content_type
            except requests.RequestException as exc:
                if attempt == max_retries:
                    raise RuntimeError(
                        f"Gagal mengunduh file {file_id} setelah "
                        f"{max_retries} percobaan: {exc}"
                    ) from exc
                time.sleep(retry_delay)
    finally:
        session.close()


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

    raw_bytes, content_type = _download_drive_bytes(
        file_id, max_retries, retry_delay, timeout
    )
    return image_bytes_to_png(raw_bytes, content_type)


def download_url_image(url: str, *, timeout: float = 15,
                       max_bytes: int = MAX_WEB_IMAGE_BYTES):
    """Download an HTTP(S) URL and return a validated, normalized PNG image.

    Google Drive share URLs use the existing confirmation-aware downloader.
    Other URLs are streamed with a size limit. The returned bytes are always
    decoded by Pillow, so misleading file extensions or Content-Type headers
    cannot cause non-image content to be inserted into a DOCX.
    """
    value = str(url or "").strip()
    parsed = urlparse(value)
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Nilai bukan tautan HTTP(S) yang valid.")

    hostname = (parsed.hostname or "").lower()
    if hostname == "drive.google.com" or hostname.endswith(".drive.google.com"):
        file_id = extract_drive_file_id(value)
        if not file_id:
            raise RuntimeError("Tautan Google Drive tidak memiliki file ID.")
        return download_drive_image(file_id, timeout=timeout)

    raw_bytes, content_type = _download_url_bytes(value, timeout, max_bytes)
    return image_bytes_to_png(raw_bytes, content_type)


def _download_url_bytes(url: str, timeout: float, max_bytes: int):
    """Download a non-Drive web resource with a strict response-size limit."""
    response = requests.get(url, stream=True, timeout=timeout)
    try:
        response.raise_for_status()
        content_length = response.headers.get("Content-Length")
        if content_length and int(content_length) > max_bytes:
            raise RuntimeError(
                f"Ukuran gambar melebihi batas {max_bytes // (1024 * 1024)} MB."
            )
        chunks = []
        total = 0
        for chunk in response.iter_content(chunk_size=64 * 1024):
            if not chunk:
                continue
            total += len(chunk)
            if total > max_bytes:
                raise RuntimeError(
                    f"Ukuran gambar melebihi batas {max_bytes // (1024 * 1024)} MB."
                )
            chunks.append(chunk)
        return b"".join(chunks), response.headers.get("Content-Type", "")
    finally:
        response.close()


def download_url_evidence(url: str, *, timeout: float = 15,
                          max_bytes: int = MAX_WEB_IMAGE_BYTES):
    """Download an HTTP(S) image/PDF as evidence layout items."""
    value = str(url or "").strip()
    parsed = urlparse(value)
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Nilai bukan tautan HTTP(S) yang valid.")

    file_id = extract_drive_file_id(value)
    if file_id and (parsed.hostname or "").lower().endswith("drive.google.com"):
        return download_drive_evidence(file_id, timeout=timeout)

    raw_bytes, content_type = _download_url_bytes(value, timeout, max_bytes)
    is_pdf = (
        raw_bytes.lstrip().startswith(b"%PDF-")
        or "application/pdf" in content_type.lower()
    )
    if is_pdf:
        return [("pdf_page", stream, size)
                for stream, size in pdf_bytes_to_png_pages(raw_bytes)]

    stream, image = image_bytes_to_png(raw_bytes, content_type)
    try:
        size = image.size
    finally:
        image.close()
    return [("image", stream, size)]


def download_drive_evidence(
    file_id: str,
    *,
    max_retries: int = 3,
    retry_delay: float = 2,
    timeout: float = 30,
):
    """Unduh bukti Drive sebagai satu gambar atau seluruh halaman PDF.

    Return value berupa daftar ``(kind, stream, (width, height))``. ``kind``
    bernilai ``image`` atau ``pdf_page`` sehingga pemanggil dapat memaksa setiap
    halaman PDF ke halaman Word tersendiri.
    """
    if not HAS_PIL:
        raise RuntimeError("Pillow tidak terinstal.")
    raw_bytes, content_type = _download_drive_bytes(
        file_id, max_retries, retry_delay, timeout
    )
    is_pdf = raw_bytes.lstrip().startswith(b"%PDF-") or "application/pdf" in content_type.lower()
    if is_pdf:
        return [("pdf_page", stream, size)
                for stream, size in pdf_bytes_to_png_pages(raw_bytes)]

    stream, image = image_bytes_to_png(raw_bytes, content_type)
    try:
        size = image.size
    finally:
        image.close()
    return [("image", stream, size)]