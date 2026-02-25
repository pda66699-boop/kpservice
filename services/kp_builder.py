from __future__ import annotations

import shutil
import unicodedata
from pathlib import Path
from typing import Any, Mapping

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream

from .docx_service import generate_docx
from .merge_service import merge_pdfs
from .pdf_service import convert_docx_to_pdf, convert_excel_to_pdf, convert_rtf_to_pdf


def _normalize_whitespace(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    normalized = normalized.replace("\u00A0", " ")
    normalized = normalized.replace("\u200B", "").replace("\u200C", "").replace("\u200D", "").replace("\uFEFF", "")
    return normalized


def _text_operands_have_visible_content(operands: list[Any]) -> bool:
    # Operators Tj/TJ/'/" may carry strings, arrays (for TJ), and positioning numbers.
    for operand in operands:
        if isinstance(operand, list):
            for nested in operand:
                if isinstance(nested, str) and _normalize_whitespace(nested).strip():
                    return True
        elif isinstance(operand, str):
            if _normalize_whitespace(operand).strip():
                return True
    return False


def _split_cover_and_rest(
    form_pdf_path: str | Path,
    cover_output_path: str | Path,
    rest_output_path: str | Path,
) -> tuple[Path, Path | None]:
    form_pdf = Path(form_pdf_path)
    cover_output = Path(cover_output_path)
    rest_output = Path(rest_output_path)

    reader = PdfReader(str(form_pdf))
    total_pages = len(reader.pages)

    if total_pages < 1:
        raise ValueError("Form PDF has no pages")

    cover_output.parent.mkdir(parents=True, exist_ok=True)

    cover_writer = PdfWriter()
    cover_writer.add_page(reader.pages[0])
    with cover_output.open("wb") as cover_file:
        cover_writer.write(cover_file)

    if total_pages == 1:
        return cover_output, None

    rest_writer = PdfWriter()
    for page in reader.pages[1:]:
        rest_writer.add_page(page)
    with rest_output.open("wb") as rest_file:
        rest_writer.write(rest_file)

    return cover_output, rest_output


def _is_page_blank(page, reader: PdfReader, min_content_bytes: int = 20) -> bool:
    text = (page.extract_text() or "").strip()
    if text:
        return False

    contents = page.get_contents()
    if contents is None:
        return True

    # Parse operators to detect pages with only clipping/state setup and no visible output.
    try:
        content_stream = ContentStream(contents, reader)
        path_paint_ops = {"S", "s", "f", "F", "f*", "B", "B*", "b", "b*"}
        text_show_ops = {"Tj", "TJ", "'", "\""}
        always_visual_ops = {"Do", "BI", "ID", "EI", "sh"}
        non_visual_ops = {
            # Graphics state / marked content.
            "q", "Q", "cm", "w", "J", "j", "M", "d", "i", "gs",
            "RG", "rg", "G", "g", "K", "k", "CS", "cs", "SC", "sc", "SCN", "scn",
            "BMC", "BDC", "EMC", "MP", "DP",
            # Path construction / clipping without paint.
            "m", "l", "c", "v", "y", "h", "re", "n", "W", "W*",
            # Text state/positioning without showing text.
            "BT", "ET", "Tf", "Td", "TD", "Tm", "Tr", "Ts", "TL", "Tc", "Tw", "Tz", "T*",
        }
        has_meaningful_visual = False
        artifact_depth = 0
        for operands, op_raw in content_stream.operations:
            op = op_raw.decode("latin1", "ignore") if isinstance(op_raw, bytes) else str(op_raw)

            if op in {"BMC", "BDC"}:
                tag = operands[0] if operands else None
                if str(tag) == "/Artifact":
                    artifact_depth += 1
                continue
            if op == "EMC":
                if artifact_depth > 0:
                    artifact_depth -= 1
                continue

            if op in text_show_ops:
                if _text_operands_have_visible_content(operands):
                    has_meaningful_visual = True
                    break
                continue

            if op in always_visual_ops:
                has_meaningful_visual = True
                break

            # Filled/stroked rectangles in Artifact blocks are usually layout noise
            # on blank trailing pages after conversion.
            if op in path_paint_ops:
                if artifact_depth == 0:
                    has_meaningful_visual = True
                    break
                continue

            if op not in non_visual_ops:
                # Unknown operator: assume it may be visible to avoid deleting real content.
                has_meaningful_visual = True
                break
        if not has_meaningful_visual:
            return True
    except Exception:
        # Fallback for malformed/unsupported content streams.
        pass

    if isinstance(contents, list):
        size = sum(len(c.get_data() or b"") for c in contents)
    else:
        size = len(contents.get_data() or b"")

    return size <= min_content_bytes


def _remove_blank_pages(input_pdf: str | Path, output_pdf: str | Path) -> Path:
    source = Path(input_pdf)
    target = Path(output_pdf)

    reader = PdfReader(str(source))
    writer = PdfWriter()

    kept = 0
    for page in reader.pages:
        if not _is_page_blank(page, reader):
            writer.add_page(page)
            kept += 1

    # Safety fallback: if all pages detected blank, keep original document as-is.
    if kept == 0:
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

    with target.open("wb") as out_file:
        writer.write(out_file)

    return target


def build_kp_pdf(
    template_path: str | Path,
    data: Mapping[str, Any],
    excel_path: str | Path,
    output_dir: str | Path,
    kp_filename: str = "kp_final.pdf",
    drawings_rtf_path: str | Path | None = None,
) -> Path:
    """Build final KP PDF: cover + price + drawings(optional) + rest_of_template."""
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    form_docx = out_dir / "form_filled.docx"
    form_pdf = out_dir / "form.pdf"
    cover_pdf = out_dir / "cover.pdf"
    rest_pdf = out_dir / "rest_template.pdf"
    price_pdf = out_dir / "price.pdf"
    drawings_pdf = out_dir / "drawings.pdf"
    final_pdf = out_dir / kp_filename

    generate_docx(template_path=template_path, data=data, output_path=form_docx)

    convert_docx_to_pdf(input_path=form_docx, output_path=form_pdf)
    convert_excel_to_pdf(input_path=excel_path, output_path=price_pdf)

    _, rest_template_pdf = _split_cover_and_rest(
        form_pdf_path=form_pdf,
        cover_output_path=cover_pdf,
        rest_output_path=rest_pdf,
    )

    merge_order: list[str] = [str(cover_pdf), str(price_pdf)]

    if drawings_rtf_path is not None:
        drawings_source = Path(drawings_rtf_path)
        if drawings_source.suffix.lower() == ".pdf":
            if not drawings_source.exists():
                raise FileNotFoundError(f"Drawings PDF not found: {drawings_source}")
            if drawings_source != drawings_pdf:
                shutil.copy2(drawings_source, drawings_pdf)
        else:
            convert_rtf_to_pdf(input_path=drawings_source, output_path=drawings_pdf)
        merge_order.append(str(drawings_pdf))

    if rest_template_pdf is not None:
        merge_order.append(str(rest_template_pdf))

    merged_pdf = merge_pdfs(files=merge_order, output_path=str(final_pdf))
    return _remove_blank_pages(input_pdf=merged_pdf, output_pdf=final_pdf)
