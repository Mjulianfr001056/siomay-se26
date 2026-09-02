"""Tests for shared JPEG/PNG/HEIC normalization."""

import io
import unittest

from docx import Document
from PIL import Image

from utils.images import HAS_HEIF, image_bytes_to_png


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


if __name__ == "__main__":
    unittest.main()