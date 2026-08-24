"""
PDF → Accessible Word Document Converter using Claude Opus API

Usage:
    python pdf_to_word.py <input.pdf> [output.docx]

Requirements:
    pip install -r requirements.txt

Set your API key:
    set ANTHROPIC_API_KEY=sk-ant-...

Implementation lives in the pdf_accessibility/ package alongside this file —
see pdf_accessibility/__init__.py for the module layout.
"""

from pdf_accessibility import main

if __name__ == "__main__":
    main()
