from __future__ import annotations

import subprocess
from shutil import which
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import quote


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
    # Force export of each sheet as a single PDF page to avoid table splitting.
    calc_filter = 'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}'
    return _convert_to_pdf(source, output_path, timeout=timeout, convert_filter=calc_filter)


def convert_rtf_to_pdf(input_path: str | Path, output_path: str | Path, timeout: int = 120) -> Path:
    """Convert RTF file to PDF using LibreOffice headless."""
    source = Path(input_path)
    if source.suffix.lower() != ".rtf":
        raise ValueError(f"Expected .rtf file, got: {source.suffix}")
    return _convert_to_pdf(source, output_path, timeout=timeout)
