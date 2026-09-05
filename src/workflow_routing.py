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
from src.document_generator import (
    custom_template_placeholders,
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

    BAPP, SPP and BAST custom placeholders must have exact, case-sensitive
    columns in their designated row-data sheet.
    """
    validator = get_input_validator(document)
    if validator is None:
        return False, [f"Validator tidak tersedia untuk dokumen: {document.id if document else '-'}"], {}
    custom_sheet = None
    if document.id.startswith("bapp_"):
        custom_sheet = "input"
    elif document.id.startswith("spp_") or document.id.startswith("bast_"):
        custom_sheet = "data_mitra"
    custom = []
    if custom_sheet and template_path and document.builtin_template_path:
        custom = custom_template_placeholders(
            template_path, document.builtin_template_path
        )
    if custom_sheet:
        return validator(file_path, custom_fields=custom)
    ok, errors, dfs = validator(file_path)
    return ok, errors, dfs