"""
Workflow configuration for the document generator wizard.

Catalog of documents, their grouping, expected built-in template filenames,
input-format references, and helpers to resolve them.
"""
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(PROJECT_ROOT, "template")
INPUT_DIR = os.path.join(PROJECT_ROOT, "input")

# Format data masukan untuk grup Lampiran SPK
INPUT_LAMPIRAN_SPK = os.path.join(INPUT_DIR, "00_input_lampiran_spk.xlsx")
INPUT_BAPP_T1_PML = os.path.join(INPUT_DIR, "01_input_bapp_t1_pml.xlsx")
INPUT_BAPP_T1_PPL = os.path.join(INPUT_DIR, "01_input_bapp_t1_ppl.xlsx")
INPUT_BAPP_T2_PML = os.path.join(INPUT_DIR, "03_input_bapp_pml_t2.xlsx")
INPUT_BAPP_T2_PPL = os.path.join(INPUT_DIR, "03_input_bapp_ppl_t2.xlsx")


class DocumentType:
    def __init__(self, doc_id: str, label: str, group: str, prefix: str,
                 template_filename: str, description: str = "",
                 implemented: bool = True):
        self.id = doc_id
        self.label = label
        self.group = group
        self.prefix = prefix                    # used for output filenames
        self.template_filename = template_filename
        self.description = description
        self.implemented = implemented          # False = "segera hadir"

    @property
    def builtin_template_path(self):
        """Absolute path to the bundled template if it exists, else None."""
        path = os.path.join(TEMPLATE_DIR, self.template_filename)
        return path if os.path.isfile(path) else None

    @property
    def kind(self):
        """PPL/PML varian dari dokumen: 'ppl' | 'pml' | None."""
        if "_ppl" in self.id:
            return "ppl"
        if "_pml" in self.id:
            return "pml"
        return None

    @property
    def input_template_path(self):
        """Path template Excel input bawaan untuk grup ini, atau None."""
        if self.group == "Lampiran SPK":
            path = INPUT_LAMPIRAN_SPK
        elif self.id == "bapp_pml_t1":
            path = INPUT_BAPP_T1_PML
        elif self.id == "bapp_ppl_t1":
            path = INPUT_BAPP_T1_PPL
        elif self.id == "bapp_pml_t2":
            path = INPUT_BAPP_T2_PML
        elif self.id == "bapp_ppl_t2":
            path = INPUT_BAPP_T2_PPL
        else:
            path = None
        return path if path and os.path.isfile(path) else None


DOCUMENT_TYPES = [
    DocumentType(
        "lampiran_spk_ppl", "Lampiran SPK PPL", "Lampiran SPK",
        "Lampiran_SPK_PPL", "00. Template Lampiran SPK PPL.docx",
        "Lampiran kontrak kerja untuk Petugas Lapangan (PPL)",
    ),
    DocumentType(
        "lampiran_spk_pml", "Lampiran SPK PML", "Lampiran SPK",
        "Lampiran_SPK_PML", "00. Template Lampiran SPK PML.docx",
        "Lampiran kontrak kerja untuk Petugas Pemeriksa Lapangan (PML)",
    ),
    # ── Grup berikutnya: belum diimplementasikan ──────────────────────
    DocumentType(
        "bapp_ppl_t1", "BAPP PPL Termin 1", "BAPP Termin 1",
        "BAPP_PPL_Termin1", "01. Template BAPP T1 PPL.docx",
        "Berita Acara Pemeriksaan Hasil Pekerjaan untuk PPL",
    ),
    DocumentType(
        "bapp_pml_t1", "BAPP PML Termin 1", "BAPP Termin 1",
        "BAPP_PML_Termin1", "01. Template BAPP T1 PML.docx",
        "Berita Acara Pemeriksaan Hasil Pekerjaan untuk PML",
    ),
    DocumentType(
        "spp_ppl", "SPP PPL", "SPP",
        "SPP_PPL", "",
        "Surat Permintaan Pembayaran untuk PPL",
        implemented=False,
    ),
    DocumentType(
        "spp_pml", "SPP PML", "SPP",
        "SPP_PML", "",
        "Surat Permintaan Pembayaran untuk PML",
        implemented=False,
    ),
    DocumentType(
        "bapp_ppl_t2", "BAPP PPL Termin 2", "BAPP Termin 2",
        "BAPP_PPL_Termin2", "03. Template BAPP T2 PPL .docx",
        "Berita Acara Pemeriksaan Hasil Pekerjaan untuk PPL (Termin II)",
    ),
    DocumentType(
        "bapp_pml_t2", "BAPP PML Termin 2", "BAPP Termin 2",
        "BAPP_PML_Termin2", "03. Template BAPP T2 PML.docx",
        "Berita Acara Pemeriksaan Hasil Pekerjaan untuk PML (Termin II)",
    ),
    DocumentType(
        "bast_ppl", "BAST PPL", "BAST",
        "BAST_PPL", "",
        "Berita Acara Serah Terima untuk PPL",
        implemented=False,
    ),
    DocumentType(
        "bast_pml", "BAST PML", "BAST",
        "BAST_PML", "",
        "Berita Acara Serah Terima untuk PML",
        implemented=False,
    ),
]

# Preserve catalog order for grouping in the UI
GROUP_ORDER = []
for _d in DOCUMENT_TYPES:
    if _d.group not in GROUP_ORDER:
        GROUP_ORDER.append(_d.group)


def get_document_by_id(doc_id: str):
    for d in DOCUMENT_TYPES:
        if d.id == doc_id:
            return d
    return None


def documents_by_group(group: str):
    return [d for d in DOCUMENT_TYPES if d.group == group]
