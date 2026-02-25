from __future__ import annotations

import argparse
from pathlib import Path

from services.kp_builder import build_kp_pdf


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build KP PDF from template, excel and drawings")
    parser.add_argument("--template", required=True, help="Path to .docx template")
    parser.add_argument("--excel", required=True, help="Path to price file (.xlsx/.xls/.xlsm)")
    parser.add_argument("--drawings", required=False, help="Path to drawings file (.rtf or .pdf)")
    parser.add_argument("--client-name", required=True, help="Client name")
    parser.add_argument("--kp-number", required=True, help="KP number")
    parser.add_argument("--manager-phone", required=True, help="Manager phone")
    parser.add_argument("--output-dir", default="artifacts/manual", help="Output directory")
    parser.add_argument("--output-name", default="kp_final.pdf", help="Output PDF filename")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    final_pdf = build_kp_pdf(
        template_path=Path(args.template),
        data={
            "client_name": args.client_name,
            "kp_number": args.kp_number,
            "manager_phone": args.manager_phone,
        },
        excel_path=Path(args.excel),
        drawings_rtf_path=Path(args.drawings) if args.drawings else None,
        output_dir=Path(args.output_dir),
        kp_filename=args.output_name,
    )

    print(final_pdf)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
