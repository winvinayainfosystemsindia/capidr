"""Check: Navigation via bookmarks for key sections.

WCAG 2.2 SC: 2.4.1 Bypass Blocks (A), 2.4.5 Multiple Ways (AA)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_headings, get_bookmarks


class CheckNavigationBookmarks(BaseCheck):
    section = "Page Layout & Navigation"
    checklist_item = "Navigation Bookmarks"
    description = "Ensure navigation via bookmarks for key sections (if required)"
    wcag_criteria = "2.4.1 Bypass Blocks (A), 2.4.5 Multiple Ways (AA)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        headings = get_headings(doc)
        content_headings = [(idx, lvl, txt) for idx, lvl, txt in headings if lvl != 6]
        bookmarks = get_bookmarks(doc)

        if not content_headings:
            results.append(
                self.not_applicable("No content headings found for navigation check")
            )
            return results

        # Headings themselves enable Navigation Pane in Word
        results.append(
            self.pass_check(
                location="Navigation Pane",
                actual=(
                    f"{len(content_headings)} heading(s) found — "
                    "these enable Word's Navigation Pane for structural navigation"
                ),
                expected="Headings should provide navigation structure",
            )
        )

        if bookmarks:
            results.append(
                self.pass_check(
                    location="Bookmarks",
                    actual=f"{len(bookmarks)} bookmark(s) available for navigation",
                    expected="Bookmarks enhance navigation for key sections",
                )
            )

        return results
