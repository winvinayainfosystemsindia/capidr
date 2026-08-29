"""CLI entry point for the validation framework.

Usage:
    python -m validation input.docx
    python -m validation input.docx -o report.xlsx
    python -m validation input.docx --checks "Document Setup & Structure" --verbose
"""

import argparse
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        prog="validation",
        description="WCAG 2.2 Remediated Document Validation Framework — "
        "validates a .docx file against the accessibility checklist "
        "and produces an Excel report.",
    )
    parser.add_argument(
        "input",
        help="Path to the remediated .docx file to validate.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Path or filename for the output Excel report. "
        "Defaults to <input_folder>/<input_name>_validation_report.xlsx.",
    )
    parser.add_argument(
        "--checks",
        nargs="*",
        default=None,
        help="Run only specific sections (e.g., 'Document Setup & Structure').",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print detailed progress to console.",
    )

    args = parser.parse_args()

    # Validate input
    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"[ERROR] File not found: {input_path}", file=sys.stderr)
        sys.exit(1)
    if input_path.suffix.lower() != ".docx":
        print(f"[ERROR] Expected a .docx file, got: {input_path.suffix}", file=sys.stderr)
        sys.exit(1)

    # Import here to avoid circular imports
    from .runner import run_validation

    try:
        run_validation(
            doc_path=str(input_path),
            output_path=args.output,
            sections=args.checks,
            verbose=args.verbose,
        )
    except Exception as exc:
        print(f"[ERROR] Validation failed: {exc}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
