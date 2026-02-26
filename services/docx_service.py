from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from docxtpl import DocxTemplate


def _replace_text_in_paragraphs(paragraphs: list[Any], old: str, new: str) -> None:
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)


def _replace_text_in_tables(tables: list[Any], old: str, new: str) -> None:
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                _replace_text_in_paragraphs(cell.paragraphs, old, new)
                _replace_text_in_tables(cell.tables, old, new)


def _normalize_symbols_for_pdf(document: DocxTemplate) -> None:
    # Emoji glyphs (e.g. 📞) often break on server LibreOffice without emoji fonts.
    # Replace with broadly supported Unicode fallback before DOCX->PDF conversion.
    doc = getattr(document, "docx", None)
    if doc is None:
        return

    _replace_text_in_paragraphs(doc.paragraphs, "📞", "☎")
    _replace_text_in_tables(doc.tables, "📞", "☎")

    for section in doc.sections:
        _replace_text_in_paragraphs(section.header.paragraphs, "📞", "☎")
        _replace_text_in_tables(section.header.tables, "📞", "☎")
        _replace_text_in_paragraphs(section.footer.paragraphs, "📞", "☎")
        _replace_text_in_tables(section.footer.tables, "📞", "☎")


def generate_docx(template_path: str | Path, data: Mapping[str, Any], output_path: str | Path) -> Path:
    """Render a DOCX template with docxtpl and save it to output_path."""
    if not isinstance(data, Mapping):
        raise TypeError("data must be a mapping")

    template = Path(template_path)
    output = Path(output_path)

    if not template.exists():
        raise FileNotFoundError(f"Template file not found: {template}")

    output.parent.mkdir(parents=True, exist_ok=True)

    document = DocxTemplate(str(template))
    document.render(dict(data))
    _normalize_symbols_for_pdf(document)
    document.save(str(output))

    return output
