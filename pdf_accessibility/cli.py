"""Argument parsing and program entry point."""

import argparse
import json
import os
import sys
from pathlib import Path

from .claude_client import upload_pdf_and_get_response
from .docx_builder import build_docx
from .pdf_images import extract_images_from_pdf
from .verify import verify_document


def load_env_file():
    """Load variables from .env file if present."""
    env_path = Path(".env")
    if env_path.exists():
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip("'\"")
                    if key and not os.environ.get(key):
                        os.environ[key] = value


def main():
    load_env_file()
    parser = argparse.ArgumentParser(
        description="Convert PDF to accessible Word document using Claude Opus API"
    )
    parser.add_argument("pdf_path", help="Path to the input PDF file")
    parser.add_argument(
        "output_path",
        nargs="?",
        default=None,
        help="Path for the output .docx file (default: same name as PDF with .docx extension)",
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="Anthropic API key (or set ANTHROPIC_API_KEY env var)",
    )
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="Only output the JSON structure (don't build .docx)",
    )
    parser.add_argument(
        "--from-json",
        default=None,
        help="Build .docx from a previously saved JSON file (skip API call)",
    )
    parser.add_argument(
        "--save-json",
        action="store_true",
        help="Save the intermediate JSON to a file alongside the .docx",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        default=True,
        help="Verify the generated document structure (default: True)",
    )

    args = parser.parse_args()

    # Determine output path
    if args.output_path is None:
        args.output_path = str(Path(args.pdf_path).with_suffix(".docx"))

    # Get API key
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")

    if args.from_json:
        # Build from existing JSON
        print(f"[INFO] Loading JSON from: {args.from_json}")
        with open(args.from_json, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        if not api_key:
            print("ERROR: No API key provided.")
            print("Set ANTHROPIC_API_KEY environment variable or use --api-key flag.")
            sys.exit(1)

        # Call Claude API
        data = upload_pdf_and_get_response(args.pdf_path, api_key)

    if args.json_only:
        print(json.dumps(data, indent=2, ensure_ascii=False))
        return

    # Save intermediate JSON if requested
    if args.save_json:
        json_path = str(Path(args.output_path).with_suffix(".json"))
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"[INFO] JSON saved to: {json_path}")

    # Extract images from the PDF
    extracted_images = {}
    if not args.from_json:
        extracted_images = extract_images_from_pdf(args.pdf_path)

    # Build the Word document
    print(f"[INFO] Building Word document...")
    build_docx(data, args.output_path, extracted_images)

    # Verify
    if args.verify:
        verify_document(args.output_path)
