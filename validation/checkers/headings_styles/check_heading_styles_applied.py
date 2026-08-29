"""Check: All headings use built-in Heading 1–6 styles.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A), 2.4.6 Headings and Labels (AA)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import looks_like_heading, get_nearest_page_marker


class CheckHeadingStylesApplied(BaseCheck):
    section = "Headings & Styles"
    checklist_item = "Heading Styles Applied"
    description = (
        "Apply correct heading styles (Heading 1, 2, 3, etc.) to all headings"
    )
    wcag_criteria = "1.3.1 Info and Relationships (A), 2.4.6 Headings and Labels (AA)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        fake_headings = []

        for idx, para in enumerate(doc.paragraphs):
            if looks_like_heading(para):
                page = get_nearest_page_marker(doc, idx)
                fake_headings.append({
                    "index": idx,
                    "text": para.text.strip()[:60],
                    "page": page,
                })

        if not fake_headings:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="All headings use built-in Heading styles",
                    expected="All headings should use Heading 1-6 styles, not manual bold/font-size",
                )
            )
        else:
            for fh in fake_headings:
                results.append(
                    self.fail_check(
                        reason="Paragraph appears to be a heading but uses manual formatting instead of a Heading style",
                        location=f"Paragraph {fh['index'] + 1} (near {fh['page']}), Text: \"{fh['text']}\"",
                        expected="Apply a built-in Heading style (Heading 1-6)",
                        actual="Manual bold/large-font formatting used instead of Heading style",
                    )
                )

        return results
