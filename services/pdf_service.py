from __future__ import annotations

import subprocess
from shutil import which
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins


class PdfConversionError(RuntimeError):
    """Raised when LibreOffice conversion fails."""


def _resolve_soffice_binary() -> str:
    from_path = which("soffice")
    if from_path:
        return from_path

    macos_default = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if macos_default.exists():
        return str(macos_default)

    raise PdfConversionError("LibreOffice (soffice) is not installed or not in PATH")


def _file_uri(path: Path) -> str:
    return f"file://{quote(str(path.resolve()))}"


def _prepare_excel_for_a4(input_path: Path, temp_dir: Path) -> Path:
    prepared = temp_dir / f"{input_path.stem}_a4{input_path.suffix.lower()}"

    workbook = load_workbook(filename=str(input_path), data_only=False, keep_vba=input_path.suffix.lower() == ".xlsm")
    for sheet in workbook.worksheets:
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 1
        sheet.page_setup.scale = None
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5, header=0.3, footer=0.3)
        sheet.print_options.horizontalCentered = True
        sheet.print_options.verticalCentered = False

    workbook.save(str(prepared))
    return prepared


def _convert_to_pdf(
    input_path: str | Path,
    output_path: str | Path,
    timeout: int = 120,
    convert_filter: str | None = None,
) -> Path:
    source = Path(input_path)
    target = Path(output_path)

    if not source.exists():
        raise FileNotFoundError(f"Input file not found: {source}")

    target.parent.mkdir(parents=True, exist_ok=True)

    soffice_bin = _resolve_soffice_binary()

    with TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        convert_to = convert_filter or "pdf"
        command = [
            soffice_bin,
            f"-env:UserInstallation={_file_uri(Path(profile_dir))}",
            "--headless",
            "--nologo",
            "--nodefault",
            "--norestore",
            "--convert-to",
            convert_to,
            "--outdir",
            str(target.parent),
            str(source),
        ]

        try:
            result = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except FileNotFoundError as exc:
            raise PdfConversionError("LibreOffice binary not found") from exc
        except subprocess.TimeoutExpired as exc:
            raise PdfConversionError(f"Conversion timed out for file: {source}") from exc

    converted_default = target.parent / f"{source.stem}.pdf"

    if result.returncode != 0:
        raise PdfConversionError(
            "LibreOffice conversion failed: "
            f"code={result.returncode}; stderr={result.stderr.strip()}"
        )

    if not converted_default.exists():
        raise PdfConversionError(f"Converted file was not created: {converted_default}")

    if converted_default != target:
        converted_default.replace(target)

    return target


def convert_docx_to_pdf(input_path: str | Path, output_path: str | Path, timeout: int = 120) -> Path:
    """Convert DOCX file to PDF using LibreOffice headless."""
    source = Path(input_path)
    if source.suffix.lower() != ".docx":
        raise ValueError(f"Expected .docx file, got: {source.suffix}")
    return _convert_to_pdf(source, output_path, timeout=timeout)


def convert_excel_to_pdf(input_path: str | Path, output_path: str | Path, timeout: int = 120) -> Path:
    """Convert Excel file (.xlsx/.xls/.xlsm) to PDF using LibreOffice headless."""
    source = Path(input_path)
    if source.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
        raise ValueError(f"Expected Excel file (.xlsx/.xls/.xlsm), got: {source.suffix}")

    # Keep A4 and fit table inside printable area.
    calc_filter = 'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}'
    if source.suffix.lower() == ".xls":
        return _convert_to_pdf(source, output_path, timeout=timeout, convert_filter=calc_filter)

    with TemporaryDirectory(prefix="excel_a4_") as temp_dir:
        prepared_excel = _prepare_excel_for_a4(source, Path(temp_dir))
        return _convert_to_pdf(prepared_excel, output_path, timeout=timeout, convert_filter=calc_filter)


def convert_rtf_to_pdf(input_path: str | Path, output_path: str | Path, timeout: int = 120) -> Path:
    """Convert RTF file to PDF using LibreOffice headless."""
    source = Path(input_path)
    if source.suffix.lower() != ".rtf":
        raise ValueError(f"Expected .rtf file, got: {source.suffix}")
    return _convert_to_pdf(source, output_path, timeout=timeout)
