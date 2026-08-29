"""Check: Reading order and navigation pane verification.

WCAG 2.2 SC: 1.3.2 Meaningful Sequence (A), 2.4.3 Focus Order (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_headings, get_text_boxes


class CheckReadingOrderNav(BaseCheck):
    section = "Final QA & Accessibility Check"
    checklist_item = "Reading Order & Navigation"
    description = "Check reading order using navigation pane and accessibility checker"
    wcag_criteria = "1.3.2 Meaningful Sequence (A), 2.4.3 Focus Order (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        headings = get_headings(doc)
        text_boxes = get_text_boxes(doc)
        content_headings = [(idx, lvl, txt) for idx, lvl, txt in headings if lvl != 6]

        # Check 1: Navigation pane usability (headings exist)
        if content_headings:
            results.append(
                self.pass_check(
                    location="Navigation Pane",
                    actual=(
                        f"{len(content_headings)} content heading(s) provide "
                        "Navigation Pane structure for keyboard navigation"
                    ),
                    expected="Document should have headings for Navigation Pane",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="No content headings found — Navigation Pane will be empty",
                    location="Navigation Pane",
                    expected="At least one heading for Navigation Pane navigation",
                    actual="0 content headings (excluding page markers)",
                )
            )

        # Check 2: No text boxes disrupting keyboard navigation
        if not text_boxes:
            results.append(
                self.pass_check(
                    location="Keyboard navigation",
                    actual="No text boxes found — keyboard tab order follows document flow",
                    expected="Content should follow logical tab/keyboard order",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason=f"{len(text_boxes)} text box(es) may disrupt keyboard navigation order",
                    location="Keyboard navigation",
                    expected="No floating text boxes that break keyboard tab order",
                    actual=f"{len(text_boxes)} text box(es) found",
                )
            )

        return results
