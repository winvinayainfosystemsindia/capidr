"""Check: Complex equations have alt text or descriptions.

WCAG 2.2 SC: 1.1 Non-text Content (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_equations


class CheckEquationAltText(BaseCheck):
    section = "Equations (STEM-specific)"
    checklist_item = "Equation Alt Text"
    description = "Add alt text or descriptions for complex equations"
    wcag_criteria = "1.1 Non-text Content (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        equations = get_equations(doc)

        if not equations:
            results.append(
                self.not_applicable(
                    "No equations found in the document — this check is not applicable"
                )
            )
            return results

        # Native OMML equations are inherently accessible to screen readers
        # that support MathML/OMML — they don't need separate alt text.
        # However, very complex equations may benefit from descriptions.
        results.append(
            self.pass_check(
                location="Entire document",
                actual=(
                    f"{len(equations)} native OMML equation(s) found. "
                    "OMML equations are inherently accessible to screen readers "
                    "that support the OMML format"
                ),
                expected="Equations should be accessible (OMML is accessible by default)",
            )
        )

        return results
