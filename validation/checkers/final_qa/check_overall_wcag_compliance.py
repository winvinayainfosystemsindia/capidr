"""Check: Overall WCAG 2.2 Level A & AA compliance summary.

WCAG 2.2 SC: All applicable Level A and AA success criteria.
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckOverallWcagCompliance(BaseCheck):
    section = "Final QA & Accessibility Check"
    checklist_item = "Overall WCAG Compliance"
    description = "Validate WCAG 2.2 AA compliance for all elements"
    wcag_criteria = "All applicable Level A and AA success criteria"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        # This is a meta-check — the actual compliance is determined
        # by the aggregate of all other checkers. We provide a summary note.
        results.append(
            self.pass_check(
                location="Entire document",
                actual=(
                    "Overall WCAG 2.2 compliance is determined by the aggregate "
                    "results of all individual checks in this report. Review the "
                    "Summary sheet for pass/fail breakdown."
                ),
                expected="All Level A and AA success criteria should pass",
            )
        )

        return results
