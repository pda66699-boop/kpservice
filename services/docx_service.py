from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from docxtpl import DocxTemplate


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
    document.save(str(output))

    return output
