"""Shared constants used across the converter package."""

from pathlib import Path

from docx.shared import Pt

MODEL = "claude-opus-5"
FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(12)
MAX_TOKENS = 128_000  # Claude Opus 5 supports large output

# Microsoft's official MathML -> OMML stylesheet, bundled locally so equation
# rendering doesn't depend on Office being installed on the machine running
# this script.
_PACKAGE_ROOT = Path(__file__).resolve().parent.parent
MML2OMML_XSL_CANDIDATES = [
    _PACKAGE_ROOT / "resources" / "MML2OMML.XSL",
    Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
]
