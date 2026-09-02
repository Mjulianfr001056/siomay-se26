"""
DOCX template filling engine.

Populates a .docx template for one data row by replacing ``{{field_name}}``
placeholders found in body paragraphs, table cells, headers and footers.

If python-docx is unavailable, callers can fall back to copy_template().
"""
import os
import re
import shutil

try:
    from docx import Document
    HAS_DOCX = True
except Exception:  # pragma: no cover - depends on environment
    Document = None
    HAS_DOCX = False

PLACEHOLDER_RE = re.compile(r"\{\{\s*(\w+)\s*\}\}")

_EMPTY_TOKENS = {"", "nan", "none", "<na>", "nat"}


def clean_value(v) -> str:
    """Normalize a cell value into a clean display string."""
    if v is None:
        return ""
    try:
        if isinstance(v, float) and v != v:  # NaN
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in _EMPTY_TOKENS else s


def build_values(record: dict) -> dict:
    """Convert a dataframe record into {key: clean_str}, dropping empties."""
    return {
        str(k): clean_value(v)
        for k, v in record.items()
        if clean_value(v) != ""
    }


def _iter_paragraphs(doc):
    """Yield every paragraph: body, tables (incl. nested), headers, footers."""
    def iter_table_paragraphs(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs
                    yield from iter_table_paragraphs(cell.tables)

    for p in doc.paragraphs:
        yield p
    yield from iter_table_paragraphs(doc.tables)
    for section in doc.sections:
        for p in section.header.paragraphs:
            yield p
        yield from iter_table_paragraphs(section.header.tables)
        for p in section.footer.paragraphs:
            yield p
        yield from iter_table_paragraphs(section.footer.tables)


def extract_template_placeholders(template_path: str) -> set[str]:
    """Return all unique ``{{field_name}}`` keys used by a DOCX template.

    This uses the same paragraph traversal as :func:`fill_row`, so validation
    includes placeholders in document tables, headers, and footers as well as
    the body. It also works when a placeholder is split across Word runs.
    """
    if not HAS_DOCX:
        raise RuntimeError("python-docx tidak tersedia")

    doc = Document(template_path)
    placeholders = set()
    for paragraph in _iter_paragraphs(doc):
        placeholders.update(PLACEHOLDER_RE.findall(paragraph.text))
    return placeholders


def validate_template_placeholders(
    template_path: str, expected_placeholders: set[str],
) -> dict:
    """Compare a template's placeholders with the expected placeholder set.

    Returns ``is_valid``, plus sorted ``missing`` and ``unexpected`` keys.
    A valid edited template must retain every placeholder downloaded in Step 2
    and must not introduce fields that the selected document does not expect.
    """
    actual = extract_template_placeholders(template_path)
    expected = set(expected_placeholders)
    return {
        "is_valid": actual == expected,
        "missing": sorted(expected - actual),
        "unexpected": sorted(actual - expected),
    }


def _replace_in_paragraph(para, values: dict):
    """
    Replace placeholders in one paragraph. Returns list of keys actually
    filled (non-empty value available).
    """
    text = para.text
    if "{{" not in text:
        return []

    filled = []

    def repl(match):
        key = match.group(1)
        val = values.get(key)
        if val:  # non-empty string
            filled.append(key)
            return val
        return match.group(0)  # keep token so it is reported unresolved

    new_text = PLACEHOLDER_RE.sub(repl, text)
    if new_text != text:
        runs = para.runs
        if runs:
            runs[0].text = new_text          # keep first-run formatting
            for r in runs[1:]:
                r.text = ""
        else:
            para.text = new_text
    return filled


def fill_row(template_path: str, values: dict, out_path: str) -> dict:
    """
    Fill one document. Returns:
      {"filled": [keys substituted], "unresolved": [tokens left in doc]}
    """
    if not HAS_DOCX:
        raise RuntimeError("python-docx tidak tersedia")

    doc = Document(template_path)
    paragraphs = list(_iter_paragraphs(doc))

    filled = set()
    for p in paragraphs:
        filled.update(_replace_in_paragraph(p, values))

    remaining = set()
    for p in paragraphs:
        remaining.update(PLACEHOLDER_RE.findall(p.text))

    unresolved = sorted(remaining - filled)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)
    return {"filled": sorted(filled), "unresolved": unresolved}


def copy_template(template_path: str, out_path: str) -> str:
    """Fallback: copy template as-is (used when python-docx is missing)."""
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    shutil.copyfile(template_path, out_path)
    return out_path


def slugify(name: str, fallback: str = "tanpa_nama") -> str:
    """Filesystem-safe fragment from a person/unit name."""
    s = "".join(ch if ch.isalnum() else "_" for ch in str(name).strip())
    s = re.sub(r"_+", "_", s).strip("_")
    return s[:40] if s else fallback
