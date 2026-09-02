"""Tests for BAPP Termin 2 evidence-image layout modes."""

import io
import unittest
from unittest.mock import patch

from docx import Document
from PIL import Image

from src import bapp_pml_t2, bapp_ppl_t2


MODULES = (bapp_ppl_t2, bapp_pml_t2)


def _downloaded_image(_file_id):
    image = Image.new("RGB", (1200, 800), color=(40, 100, 180))
    stream = io.BytesIO()
    image.save(stream, format="PNG")
    stream.seek(0)
    return stream, image


def _document_with_placeholder(module):
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("Sebelum ")
    paragraph.add_run(module.BUKTI_PLACEHOLDER[:14])
    paragraph.add_run(module.BUKTI_PLACEHOLDER[14:] + " sesudah")
    trailing_paragraph = doc.add_paragraph("KONTEN SETELAH PLACEHOLDER")
    return doc, paragraph, trailing_paragraph


class BappTermin2ImageLayoutTests(unittest.TestCase):
    def test_bapp_sequence_number_preserves_three_digit_form(self):
        for module in MODULES:
            with self.subTest(module=module.__name__):
                self.assertEqual(module._format_no_urut_bapp_t2("001"), "001")
                # Excel commonly stores a displayed 001 as numeric 1.
                self.assertEqual(module._format_no_urut_bapp_t2(1), "001")
                self.assertEqual(module._format_no_urut_bapp_t2("21"), "021")
                self.assertEqual(module._format_no_urut_bapp_t2("AB-1"), "AB-1")

    def test_dedicated_pages_insert_every_image_without_grid(self):
        links = ",".join(
            f"https://drive.google.com/file/d/image-{number}/view"
            for number in range(1, 7)
        )

        for module in MODULES:
            with self.subTest(module=module.__name__):
                doc, placeholder_paragraph, trailing_paragraph = (
                    _document_with_placeholder(module)
                )
                with patch.object(
                    module, "_download_drive_image", side_effect=_downloaded_image
                ) as downloader:
                    count, warnings = module.insert_gdrive_images(
                        doc,
                        links,
                        image_layout=module.IMAGE_LAYOUT_DEDICATED_PAGES,
                    )

                self.assertEqual(count, 6)
                self.assertEqual(warnings, [])
                self.assertEqual(downloader.call_count, 6)
                self.assertEqual(len(doc.inline_shapes), 6)
                self.assertEqual(len(doc.tables), 0)
                for shape in doc.inline_shapes:
                    self.assertLessEqual(
                        shape.width.inches, module.DEDICATED_MAX_WIDTH_IN
                    )
                    self.assertLessEqual(
                        shape.height.inches, module.DEDICATED_MAX_HEIGHT_IN
                    )
                self.assertNotIn(module.BUKTI_PLACEHOLDER, placeholder_paragraph.text)
                self.assertEqual(
                    len(doc.element.body.xpath('.//w:br[@w:type="page"]')), 5
                )
                titles = [p.text for p in doc.paragraphs]
                self.assertIn("BUKTI DUKUNG (1/6)", titles)
                self.assertIn("BUKTI DUKUNG (6/6)", titles)
                # Gambar pertama langsung mengikuti placeholder; seluruh bukti
                # tetap disisipkan sebelum konten template sesudah placeholder.
                body_elements = list(doc.element.body)
                first_title = next(
                    p for p in doc.paragraphs if p.text == "BUKTI DUKUNG (1/6)"
                )
                last_image_paragraph = next(
                    p for p in reversed(doc.paragraphs) if p._p.xpath(".//w:drawing")
                )
                self.assertEqual(
                    body_elements.index(first_title._p),
                    body_elements.index(placeholder_paragraph._p) + 1,
                )
                self.assertLess(
                    body_elements.index(last_image_paragraph._p),
                    body_elements.index(trailing_paragraph._p),
                )

    def test_grid_keeps_five_image_limit_and_warning(self):
        links = ",".join(
            f"https://drive.google.com/file/d/image-{number}/view"
            for number in range(1, 7)
        )

        for module in MODULES:
            with self.subTest(module=module.__name__):
                doc, _, _ = _document_with_placeholder(module)
                with patch.object(
                    module, "_download_drive_image", side_effect=_downloaded_image
                ) as downloader:
                    count, warnings = module.insert_gdrive_images(
                        doc, links, image_layout=module.IMAGE_LAYOUT_GRID
                    )

                self.assertEqual(count, 5)
                self.assertEqual(downloader.call_count, 5)
                self.assertEqual(len(doc.inline_shapes), 5)
                self.assertGreater(len(doc.tables), 0)
                self.assertTrue(any("5 tautan pertama" in warning for warning in warnings))

    def test_invalid_layout_is_rejected(self):
        for module in MODULES:
            with self.subTest(module=module.__name__):
                doc, _, _ = _document_with_placeholder(module)
                with self.assertRaisesRegex(ValueError, "tata letak gambar"):
                    module.insert_gdrive_images(
                        doc, "", image_layout="unknown-layout"
                    )

    def test_empty_links_still_remove_placeholder(self):
        for module in MODULES:
            with self.subTest(module=module.__name__):
                doc, paragraph, _ = _document_with_placeholder(module)
                count, warnings = module.insert_gdrive_images(doc, "")

                self.assertEqual(count, 0)
                self.assertEqual(warnings, [])
                self.assertEqual(paragraph.text, "Sebelum  sesudah")


if __name__ == "__main__":
    unittest.main()