"""Regression tests for SPP Termin II catalog and generator integration."""
import os
import tempfile
import unittest
import zipfile
from pathlib import Path

import pandas as pd
from docx import Document

from src import spp_t2
from src.workflow import DOCUMENT_TYPES, get_document_by_id


ROOT = Path(__file__).resolve().parents[1]


def _document_xml_text(path):
    with zipfile.ZipFile(path) as archive:
        return "".join(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in archive.namelist()
            if name.startswith("word/") and name.endswith(".xml")
        )


class SppTermin2WorkflowTests(unittest.TestCase):
    def test_catalog_places_spp_termin_ii_before_bast_with_shifted_assets(self):
        ids = [doc.id for doc in DOCUMENT_TYPES]
        self.assertLess(ids.index("spp_t2_ppl"), ids.index("bast_ppl"))
        self.assertLess(ids.index("spp_t2_pml"), ids.index("bast_pml"))
        self.assertEqual(
            get_document_by_id("spp_t2_ppl").template_filename,
            "04. Template SPP T2 PPL.docx",
        )
        self.assertEqual(
            get_document_by_id("bast_ppl").template_filename,
            "05. Template BAST PPL.docx",
        )
        self.assertEqual(
            Path(get_document_by_id("bukti_terima").input_template_path).name,
            "06_input_bukti_terima_paket_internet.xlsx",
        )


class SppTermin2GeneratorTests(unittest.TestCase):
    def test_generator_replaces_template_fields_from_matching_input_columns(self):
        with tempfile.TemporaryDirectory() as directory:
            template_path = os.path.join(directory, "template.docx")
            document = Document()
            paragraph = document.add_paragraph()
            paragraph.add_run("{{nama_")
            paragraph.add_run("lengkap}}")
            document.add_paragraph("{{nik}}")
            document.add_paragraph("{{no_spk}}")
            document.add_paragraph("{{no_input_spp_t2}}")
            document.save(template_path)

            dataframe = pd.DataFrame([{
                "nik": "6304000000000001",
                "nama_lengkap": "PETUGAS UJI PPL",
                "no_spk": "SPK-002/2026",
                "no_input_spp_t2": "21",
            }])
            events = list(spp_t2.iter_generate(
                "ppl", {spp_t2.SHEET_NAME: dataframe}, template_path, directory
            ))
            output_path = next(event["path"] for event in events if event["t"] == "file")
            xml_text = _document_xml_text(output_path)

        self.assertIn("PETUGAS UJI PPL", xml_text)
        self.assertIn("6304000000000001", xml_text)
        self.assertIn("SPK-002/2026", xml_text)
        self.assertIn("21", xml_text)
        self.assertNotIn("{{", xml_text)

    def test_validation_requires_columns_used_by_template(self):
        with tempfile.TemporaryDirectory() as directory:
            template_path = os.path.join(directory, "template.docx")
            document = Document()
            document.add_paragraph("{{nik}} {{nama_lengkap}} {{no_spk}} {{no_input_spp_t2}}")
            document.save(template_path)
            input_path = os.path.join(directory, "input.xlsx")
            pd.DataFrame([{
                "nik": "6304000000000001",
                "nama_lengkap": "PETUGAS UJI",
                "no_spk": "SPK-002/2026",
            }]).to_excel(input_path, index=False, sheet_name=spp_t2.SHEET_NAME)

            ok, errors, _ = spp_t2.validate_input(input_path, template_path)

        self.assertFalse(ok)
        self.assertIn("no_input_spp_t2", errors[0])


if __name__ == "__main__":
    unittest.main()