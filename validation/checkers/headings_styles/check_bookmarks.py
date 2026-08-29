"""Check: Bookmarks exist for major sections.

WCAG 2.2 SC: 2.4.1 Bypass Blocks (A), 2.4.5 Multiple Ways (AA)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_bookmarks, get_headings


class CheckBookmarks(BaseCheck):
    section = "Headings & Styles"
    checklist_item = "Bookmarks"
    description = "Verify bookmarks for major sections (if required)"
    wcag_criteria = "2.4.1 Bypass Blocks (A), 2.4.5 Multiple Ways (AA)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        bookmarks = get_bookmarks(doc)
        headings = get_headings(doc)

        # Filter to major headings (H1-H3, excluding H6 page markers)
        major_headings = [(idx, lvl, txt) for idx, lvl, txt in headings if lvl <= 3]

        if not major_headings:
            results.append(
                self.not_applicable(
                    "No major headings (H1-H3) found; bookmarks check not applicable"
                )
            )
            return results

        if bookmarks:
            results.append(
                self.pass_check(
                    location="Document structure",
                    actual=f"{len(bookmarks)} bookmark(s) found: {', '.join(b['name'][:20] for b in bookmarks[:5])}",
                    expected="Major sections should have bookmarks for navigation",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="No bookmarks found for major sections",
                    location="Document structure",
                    expected=(
                        f"Bookmarks should exist for major sections "
                        f"({len(major_headings)} major headings found)"
                    ),
                    actual="0 bookmarks defined",
                )
            )

        return results
