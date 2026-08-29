"""Check: No unnecessary blank pages exist in the document.

WCAG 2.2 SC: No direct WCAG mapping (supports 1.3.2 Meaningful Sequence).
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_nearest_page_marker


class CheckBlankPages(BaseCheck):
    section = "Document Setup & Structure"
    checklist_item = "Blank Pages"
    description = "Remove unnecessary blank pages"
    wcag_criteria = "No direct WCAG mapping (supports 1.3.2)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        blank_page_locations = []

        # Detect sequences of empty paragraphs that span page breaks
        consecutive_empty = 0
        empty_start_idx = -1

        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            has_break = False

            # Check for page break in runs
            for run in para.runs:
                for br in run._element.findall(qn("w:br")):
                    if br.get(qn("w:type")) == "page":
                        has_break = True

            if not text and not has_break:
                if consecutive_empty == 0:
                    empty_start_idx = idx
                consecutive_empty += 1
            else:
                if has_break and not text:
                    # Page break with no content = potential blank page
                    page = get_nearest_page_marker(doc, idx)
                    blank_page_locations.append(
                        f"Paragraph {idx + 1} (near {page})"
                    )
                consecutive_empty = 0

        if not blank_page_locations:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="No unnecessary blank pages detected",
                    expected="No blank pages",
                )
            )
        else:
            for loc in blank_page_locations:
                results.append(
                    self.fail_check(
                        reason="Potential unnecessary blank page detected",
                        location=loc,
                        expected="No unnecessary blank pages",
                        actual="Empty page break found with no content",
                    )
                )

        return results
