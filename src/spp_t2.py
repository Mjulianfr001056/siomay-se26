"""Generator Surat Pernyataan Penyelesaian (SPP) Termin 2.

Termin 2 uses the same three-sheet data model and generation calculations as
Termin 1. The only schema difference is ``no_urut_spp_t2`` in ``no_spk``;
its value is written to the matching Word placeholder.
"""
import re

import pandas as pd
from docx import Document

from src import spp


SHEET_DATA_MITRA = spp.SHEET_DATA_MITRA
SHEET_NO_SPK = spp.SHEET_NO_SPK
SHEET_ALOKASI = spp.SHEET_ALOKASI
SHEET_NAME = SHEET_DATA_MITRA
COL_NO_URUT_SPP_T2 = "no_urut_spp_t2"
NUMBER_PLACEHOLDER = "no_urut_spp_t2"

REQUIRED_SCHEMA = {
    SHEET_DATA_MITRA: ["nik", "nama_lengkap", "jabatan"],
    SHEET_NO_SPK: ["nik", "no_spk", COL_NO_URUT_SPP_T2],
    SHEET_ALOKASI: ["nik_ppl", "nik_pml", "target", "capaian", "persentase"],
}

SUPPORTED_TEMPLATE_FIELDS = {
    "nik", "nama_lengkap", "no_spk", NUMBER_PLACEHOLDER,
    "jml_usaha", "jml_usaha_min", "persentase",
}


def _template_fields(template_path: str) -> set[str]:
    """Return placeholder names, including fields located inside tables."""
    doc = Document(template_path)
    text = []

    def visit_table(table):
        for row in table.rows:
            for cell in row.cells:
                text.extend(paragraph.text for paragraph in cell.paragraphs)
                for nested in cell.tables:
                    visit_table(nested)

    text.extend(paragraph.text for paragraph in doc.paragraphs)
    for table in doc.tables:
        visit_table(table)
    for section in doc.sections:
        text.extend(paragraph.text for paragraph in section.header.paragraphs)
        text.extend(paragraph.text for paragraph in section.footer.paragraphs)
    return set(re.findall(r"\{\{\s*(\w+)\s*\}\}", "\n".join(text)))


def validate_input(file_path: str, template_path: str | None = None):
    """Validate the Termin 2 three-sheet workbook and template placeholders."""
    errors = []
    dfs = {}
    try:
        workbook = pd.ExcelFile(file_path)
    except Exception as exc:
        return False, [f"Gagal membaca file Excel: {exc}"], {}

    try:
        for sheet_name, required_columns in REQUIRED_SCHEMA.items():
            if sheet_name not in workbook.sheet_names:
                errors.append(f"Sheet '{sheet_name}' tidak ditemukan.")
                continue
            dataframe = pd.read_excel(workbook, sheet_name=sheet_name, dtype=str)
            dataframe.columns = dataframe.columns.str.strip()
            dataframe = dataframe.fillna("")
            missing = [
                column for column in required_columns
                if column not in dataframe.columns
            ]
            if missing:
                errors.append(
                    f"Sheet '{sheet_name}' kekurangan kolom: "
                    + ", ".join(missing)
                )
            else:
                dfs[sheet_name] = dataframe
    finally:
        workbook.close()

    if template_path and not errors:
        unmapped = sorted(
            _template_fields(template_path) - SUPPORTED_TEMPLATE_FIELDS
        )
        if unmapped:
            errors.append(
                "Placeholder template SPP Termin 2 tidak didukung: "
                + ", ".join(unmapped)
            )
    return not errors, errors, dfs


def iter_generate(kind: str, dfs: dict, template_path: str, out_dir: str):
    """Generate Termin 2 with the shared SPP aggregation/table workflow."""
    yield from spp.iter_generate(
        kind,
        dfs,
        template_path,
        out_dir,
        number_column=COL_NO_URUT_SPP_T2,
        number_placeholder=NUMBER_PLACEHOLDER,
        termin_label="_Termin2",
    )