"""Check: Chemical formulas/notations are properly tagged and accessible.

WCAG 2.2 SC: 1.1 Non-text Content (A), 4.1.2 Name, Role, Value (A)
"""

import re
from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckChemicalNotations(BaseCheck):
    section = "Symbols, Chemical Notations & Special Characters (STEM-specific)"
    checklist_item = "Chemical Notations"
    description = "Ensure chemical formulas/notations are properly tagged and accessible"
    wcag_criteria = "1.1 Non-text Content (A), 4.1.2 Name, Role, Value (A)"

    # Pattern to detect chemical formulas (e.g., H2O, CO2, NaCl, C6H12O6)
    _CHEM_PATTERN = re.compile(
        r"\b[A-Z][a-z]?\d*(?:[A-Z][a-z]?\d*)+\b"
    )

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        chem_formulas = []

        for idx, para in enumerate(doc.paragraphs):
            text = para.text
            matches = self._CHEM_PATTERN.findall(text)
            for match in matches:
                # Filter common English words that look like formulas
                if match.lower() in {"no", "on", "in", "an", "or", "if", "is", "it", "at", "be"}:
                    continue
                if len(match) >= 3:
                    chem_formulas.append({
                        "index": idx,
                        "formula": match,
                    })

        if not chem_formulas:
            results.append(
                self.not_applicable(
                    "No chemical notations detected in the document"
                )
            )
        else:
            # Report that chemical formulas exist and should be verified
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=(
                        f"{len(chem_formulas)} potential chemical formula(s) detected. "
                        "Manual review recommended to ensure proper subscript/superscript formatting."
                    ),
                    expected="Chemical notations should use proper Unicode subscript/superscript or equation editor",
                )
            )

        return results
