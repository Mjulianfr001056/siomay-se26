"""Stable document-ID routing for Excel validation and generation.

Keep routing independent from UI labels and group names so terminology changes
cannot silently send a document type through another workflow's schema.
"""
from src import (
    bapp_pml,
    bapp_pml_t2,
    bapp_ppl,
    bapp_ppl_t2,
    bast,
    bukti_terima,
    lampiran_spk,
    spp,
    spp_t2,
)


INPUT_VALIDATORS = {
    "lampiran_spk_ppl": lampiran_spk.validate_input,
    "lampiran_spk_pml": lampiran_spk.validate_input,
    "bapp_ppl_t1": bapp_ppl.validate_input,
    "bapp_pml_t1": bapp_pml.validate_input,
    "spp_ppl": spp.validate_input,
    "spp_pml": spp.validate_input,
    "bapp_ppl_t2": bapp_ppl_t2.validate_input,
    "bapp_pml_t2": bapp_pml_t2.validate_input,
    "spp_t2_ppl": spp_t2.validate_input,
    "spp_t2_pml": spp_t2.validate_input,
    "bast_ppl": bast.validate_input,
    "bast_pml": bast.validate_input,
    "bukti_terima": bukti_terima.validate_input,
}


DOCUMENT_GENERATORS = {
    "lampiran_spk_ppl": lampiran_spk.iter_generate,
    "lampiran_spk_pml": lampiran_spk.iter_generate,
    "bapp_ppl_t1": bapp_ppl.iter_generate,
    "bapp_pml_t1": bapp_pml.iter_generate,
    "spp_ppl": spp.iter_generate,
    "spp_pml": spp.iter_generate,
    "bapp_ppl_t2": bapp_ppl_t2.iter_generate,
    "bapp_pml_t2": bapp_pml_t2.iter_generate,
    "spp_t2_ppl": spp_t2.iter_generate,
    "spp_t2_pml": spp_t2.iter_generate,
    "bast_ppl": bast.iter_generate,
    "bast_pml": bast.iter_generate,
    "bukti_terima": bukti_terima.iter_generate,
}


def get_input_validator(document):
    """Return the Excel validator registered for a catalog document."""
    return INPUT_VALIDATORS.get(document.id) if document else None


def get_document_generator(document):
    """Return the generator registered for a catalog document."""
    return DOCUMENT_GENERATORS.get(document.id) if document else None


def validate_document_input(document, file_path: str, template_path: str | None = None):
    """Validate an Excel file using stable document-ID dispatch.

    SPP Termin 2 additionally verifies that every Word placeholder has a
    matching Excel column. Other workflows have explicit/static mappings.
    """
    validator = get_input_validator(document)
    if validator is None:
        return False, [f"Validator tidak tersedia untuk dokumen: {document.id if document else '-'}"], {}
    if document.id in ("spp_t2_ppl", "spp_t2_pml"):
        return validator(file_path, template_path=template_path)
    return validator(file_path)