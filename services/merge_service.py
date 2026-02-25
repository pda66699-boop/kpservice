from __future__ import annotations

from pathlib import Path

from pypdf import PdfWriter


def merge_pdfs(files: list, output_path: str) -> Path:
    """Merge multiple PDF files into a single PDF."""
    if not files:
        raise ValueError("files list must not be empty")

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    writer = PdfWriter()
    for file_path in files:
        source = Path(file_path)
        if source.suffix.lower() != ".pdf":
            raise ValueError(f"Expected PDF file, got: {source}")
        if not source.exists():
            raise FileNotFoundError(f"PDF file not found: {source}")
        writer.append(str(source))

    with output.open("wb") as out_file:
        writer.write(out_file)
    writer.close()

    return output
