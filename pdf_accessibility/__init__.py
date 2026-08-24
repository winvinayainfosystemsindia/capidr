"""PDF -> Accessible Word Document Converter.

Package layout:
    constants.py      Shared constants (model name, fonts, stylesheet paths).
    prompts.py         The Claude system/user prompts driving extraction.
    pdf_images.py       PDF image extraction (PyMuPDF).
    claude_client.py    Claude API upload + JSON response parsing.
    equations.py        LaTeX -> MathML -> OMML equation rendering.
    rich_text.py         Hyperlink + inline-equation-aware text runs.
    docx_builder.py      Assembles the structured JSON into a .docx file.
    verify.py            Post-build structural verification report.
    cli.py               Argument parsing and program entry point.
"""

from .cli import main

__all__ = ["main"]
