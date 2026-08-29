"""Check: Logical document flow from start to end (reading order).

WCAG 2.2 SC: 1.3.2 Meaningful Sequence (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_headings, get_text_boxes


class CheckLogicalReadingOrder(BaseCheck):
    section = "Document Setup & Structure"
    checklist_item = "Logical Reading Order"
    description = "Confirm logical document flow from start to end"
    wcag_criteria = "1.3.2 Meaningful Sequence (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        issues = []

        # Check 1: Text boxes disrupt reading order
        text_boxes = get_text_boxes(doc)
        if text_boxes:
            issues.append(
                f"{len(text_boxes)} text box(es) found that may disrupt reading order"
            )

        # Check 2: Verify headings appear in logical order (structure check)
        headings = get_headings(doc)
        content_headings = [(idx, lvl, txt) for idx, lvl, txt in headings if lvl != 6]

        if content_headings:
            # First content heading should ideally be H1
            first_idx, first_level, first_text = content_headings[0]
            if first_level != 1:
                issues.append(
                    f"Document starts with Heading {first_level} instead of Heading 1 "
                    f"(Paragraph {first_idx + 1}: \"{first_text[:50]}\")"
                )

        # Check 3: Verify content exists between headings (no empty sections)
        for i, (h_idx, h_level, h_text) in enumerate(content_headings):
            if i + 1 < len(content_headings):
                next_idx = content_headings[i + 1][0]
                # Check if there's any non-empty content between headings
                has_content = False
                for p_idx in range(h_idx + 1, next_idx):
                    if p_idx < len(doc.paragraphs):
                        para = doc.paragraphs[p_idx]
                        if para.text.strip() and not para.style.name.startswith("Heading"):
                            has_content = True
                            break
                if not has_content and (next_idx - h_idx) == 1:
                    issues.append(
                        f"Empty section under \"{h_text[:40]}\" at Paragraph {h_idx + 1}"
                    )

        if not issues:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="Document follows a logical reading order",
                    expected="Logical flow: title → headings → content with no disruptions",
                )
            )
        else:
            for issue in issues:
                results.append(
                    self.fail_check(
                        reason=issue,
                        location="Document structure",
                        expected="Logical, sequential reading order without disruptions",
                        actual=issue,
                    )
                )

        return results
