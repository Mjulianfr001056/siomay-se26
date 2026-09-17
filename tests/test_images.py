"""Tests for shared JPEG/PNG/HEIC normalization."""

import io
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from docx import Document
from PIL import Image

from utils.images import (
    HAS_HEIF,
    HAS_PDF_RENDERER,
    extract_drive_folder_id,
    image_bytes_to_png,
    list_drive_folder_images,
    pdf_bytes_to_png_pages,
    resolve_drive_inputs,
)


class ImageNormalizationTests(unittest.TestCase):
    def test_jpeg_is_normalized_to_insertable_png(self):
        source = io.BytesIO()
        Image.new("RGB", (12, 8), "blue").save(source, format="JPEG")

        png_file, image = image_bytes_to_png(source.getvalue(), "image/jpeg")

        self.assertEqual(png_file.read(8), b"\x89PNG\r\n\x1a\n")
        self.assertEqual(image.size, (12, 8))
        png_file.seek(0)
        Document().add_paragraph().add_run().add_picture(png_file)

    @unittest.skipUnless(HAS_HEIF, "pillow-heif/native HEIF codec unavailable")
    def test_heic_is_decoded_to_insertable_png(self):
        source = io.BytesIO()
        Image.new("RGB", (11, 7), "green").save(source, format="HEIF")

        png_file, image = image_bytes_to_png(source.getvalue(), "image/heic")

        self.assertEqual(png_file.read(8), b"\x89PNG\r\n\x1a\n")
        self.assertEqual(image.size, (11, 7))
        png_file.seek(0)
        Document().add_paragraph().add_run().add_picture(png_file)

    def test_html_download_is_rejected_with_actionable_message(self):
        with self.assertRaisesRegex(RuntimeError, "halaman HTML"):
            image_bytes_to_png(b"<!doctype html><html></html>", "text/html")

    @unittest.skipUnless(HAS_PDF_RENDERER, "PyMuPDF unavailable")
    def test_multipage_pdf_is_rendered_to_insertable_png_pages(self):
        import pymupdf as fitz

        pdf = fitz.open()
        for text in ("Halaman satu", "Halaman dua"):
            page = pdf.new_page(width=300, height=500)
            page.insert_text((40, 60), text)
        raw_pdf = pdf.tobytes()
        pdf.close()

        pages = pdf_bytes_to_png_pages(raw_pdf)

        self.assertEqual(len(pages), 2)
        for stream, size in pages:
            self.assertEqual(stream.read(8), b"\x89PNG\r\n\x1a\n")
            self.assertGreater(size[0], 0)
            self.assertGreater(size[1], 0)
            stream.seek(0)
            Document().add_paragraph().add_run().add_picture(stream)


class DriveFolderTests(unittest.TestCase):
    def test_infers_folder_without_misclassifying_file_url(self):
        self.assertEqual(
            extract_drive_folder_id(
                "https://drive.google.com/drive/u/1/folders/folder-123?usp=sharing"
            ),
            "folder-123",
        )
        self.assertIsNone(extract_drive_folder_id(
            "https://drive.google.com/file/d/file-123/view"
        ))

    def test_folder_requires_configured_worker_endpoint(self):
        with patch.dict(os.environ, {"SIOMAY_DRIVE_FOLDER_WORKER_URL": ""}):
            with self.assertRaisesRegex(RuntimeError, "belum dikonfigurasi"):
                list_drive_folder_images("folder-12345")

    def test_folder_listing_reads_worker_url_from_dotenv(self):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {"files": []}
        requester = Mock(return_value=response)

        with tempfile.TemporaryDirectory() as directory:
            dotenv = Path(directory) / ".env"
            dotenv.write_text(
                "SIOMAY_DRIVE_FOLDER_WORKER_URL=https://dotenv-worker.example\n",
                encoding="utf-8",
            )
            with patch.dict(os.environ, {}, clear=False), patch(
                "utils.config._dotenv_paths", return_value=(dotenv,)
            ):
                os.environ.pop("SIOMAY_DRIVE_FOLDER_WORKER_URL", None)
                files = list_drive_folder_images(
                    "folder-12345", requester=requester
                )

        self.assertEqual(files, [])
        requester.assert_called_once_with(
            "https://dotenv-worker.example/v1/drive/folders/folder-12345/images",
            timeout=30,
        )

    def test_process_environment_overrides_dotenv(self):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {"files": []}
        requester = Mock(return_value=response)

        with tempfile.TemporaryDirectory() as directory:
            dotenv = Path(directory) / ".env"
            dotenv.write_text(
                "SIOMAY_DRIVE_FOLDER_WORKER_URL=https://dotenv-worker.example\n",
                encoding="utf-8",
            )
            with patch.dict(os.environ, {
                "SIOMAY_DRIVE_FOLDER_WORKER_URL": "https://environment-worker.example"
            }), patch("utils.config._dotenv_paths", return_value=(dotenv,)):
                list_drive_folder_images("folder-12345", requester=requester)

        requester.assert_called_once_with(
            "https://environment-worker.example/v1/drive/folders/folder-12345/images",
            timeout=30,
        )

    def test_folder_listing_uses_worker_without_credentials(self):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "files": [
                {"id": "image-00002", "name": "image2.png", "mimeType": "image/png"},
                {"id": "image-00010", "name": "image10.jpg", "mimeType": "image/jpeg"},
            ],
        }
        requester = Mock(return_value=response)

        files = list_drive_folder_images(
            "folder-12345",
            requester=requester,
            worker_url="https://worker.example",
        )

        self.assertEqual([item["id"] for item in files], ["image-00002", "image-00010"])
        requester.assert_called_once_with(
            "https://worker.example/v1/drive/folders/folder-12345/images",
            timeout=30,
        )

    def test_folder_listing_rejects_insecure_endpoint_and_invalid_payload(self):
        with self.assertRaisesRegex(RuntimeError, "tidak aman"):
            list_drive_folder_images(
                "folder-12345", worker_url="http://worker.example"
            )

        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "files": [{"id": "bad!", "name": "x.jpg", "mimeType": "image/jpeg"}]
        }
        with self.assertRaisesRegex(RuntimeError, "Respons layanan"):
            list_drive_folder_images(
                "folder-12345",
                requester=Mock(return_value=response),
                worker_url="https://worker.example",
            )

    def test_folder_listing_maps_rate_limit_without_leaking_response(self):
        response = Mock(status_code=429)
        response.raise_for_status.side_effect = RuntimeError("private service detail")
        with self.assertRaisesRegex(RuntimeError, "terlalu banyak permintaan") as caught:
            list_drive_folder_images(
                "folder-12345",
                requester=Mock(return_value=response),
                worker_url="https://worker.example",
            )
        self.assertNotIn("private service detail", str(caught.exception))

    def test_resolver_preserves_mixed_input_order(self):
        folder_lister = Mock(return_value=[
            {"id": "folder-1", "name": "01.jpg", "mimeType": "image/jpeg"},
            {"id": "folder-2", "name": "02.jpg", "mimeType": "image/jpeg"},
        ])
        value = (
            "https://drive.google.com/file/d/before/view,"
            "https://drive.google.com/drive/folders/photos,"
            "https://drive.google.com/file/d/after/view"
        )

        references, warnings = resolve_drive_inputs(value, folder_lister=folder_lister)

        self.assertEqual(
            [file_id for file_id, _label in references],
            ["before", "folder-1", "folder-2", "after"],
        )
        self.assertEqual(warnings, [])


if __name__ == "__main__":
    unittest.main()