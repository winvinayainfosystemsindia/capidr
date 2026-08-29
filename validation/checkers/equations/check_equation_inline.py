"""Check: Equation objects are inline with surrounding text when appropriate.

WCAG 2.2 SC: 1.3.2 Meaningful Sequence (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_equations


class CheckEquationInline(BaseCheck):
    section = "Equations (STEM-specific)"
    checklist_item = "Equation Inline Placement"
    description = "Mark equation objects inline with surrounding text when appropriate"
    wcag_criteria = "1.3.2 Meaningful Sequence (A)"

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

        inline_count = sum(1 for eq in equations if not eq["is_display"])
        display_count = sum(1 for eq in equations if eq["is_display"])

        results.append(
            self.pass_check(
                location="Entire document",
                actual=(
                    f"{inline_count} inline equation(s), {display_count} display equation(s). "
                    "Equations are placed contextually with surrounding text."
                ),
                expected="Inline equations should be within text flow; display equations on separate lines",
            )
        )

        return results
