"""Check: All content is in the main document body (not in text boxes).

WCAG 2.2 SC: 1.3.1 Info and Relationships (A), 1.3.2 Meaningful Sequence (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_text_boxes


class CheckContentInBody(BaseCheck):
    section = "Document Setup & Structure"
    checklist_item = "Content in Body"
    description = (
        "Ensure all content is in main document body (not in text boxes)"
    )
    wcag_criteria = "1.3.1 Info and Relationships (A), 1.3.2 Meaningful Sequence (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        text_boxes = get_text_boxes(doc)

        if not text_boxes:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="No text boxes found — all content is in the main body",
                    expected="All content should be in the main document body",
                )
            )
        else:
            for tb in text_boxes:
                results.append(
                    self.fail_check(
                        reason="Content found in a text box instead of the main document body",
                        location=tb["location"],
                        expected="Content should be in the main document body for proper reading order",
                        actual=f"Text box contains: \"{tb['text']}\"",
                    )
                )

        return results
