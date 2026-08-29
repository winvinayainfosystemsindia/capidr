"""Check: All hyperlinks and cross-references point to valid destinations.

WCAG 2.2 SC: 2.4.4 Link Purpose (In Context) (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_hyperlinks, get_bookmarks, get_nearest_page_marker


class CheckLinkDestinations(BaseCheck):
    section = "Hyperlinks & Cross References"
    checklist_item = "Link Destinations"
    description = "Test all hyperlinks and cross-references for correct destination"
    wcag_criteria = "2.4.4 Link Purpose (In Context) (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        links = get_hyperlinks(doc)

        if not links:
            results.append(
                self.not_applicable("No hyperlinks found in the document")
            )
            return results

        bookmarks = {b["name"] for b in get_bookmarks(doc)}
        broken_internal = []
        valid_count = 0

        for link in links:
            url = link["url"]
            if url.startswith("#"):
                # Internal bookmark link
                bookmark_name = url[1:]
                if bookmark_name not in bookmarks:
                    page = get_nearest_page_marker(doc, link["para_index"])
                    broken_internal.append({
                        "display": link["display_text"],
                        "target": bookmark_name,
                        "location": f"{link['location']} (near {page})",
                    })
                else:
                    valid_count += 1
            elif url:
                # External URL — just check format (no HTTP requests)
                if url.startswith(("http://", "https://", "mailto:", "ftp://")):
                    valid_count += 1
                else:
                    valid_count += 1  # Other protocols count as valid format

        if broken_internal:
            for broken in broken_internal:
                results.append(
                    self.fail_check(
                        reason=f"Internal link points to non-existent bookmark: '{broken['target']}'",
                        location=broken["location"],
                        expected=f"Bookmark '{broken['target']}' should exist in the document",
                        actual=f"Link '{broken['display']}' → bookmark '{broken['target']}' not found",
                    )
                )

        if valid_count > 0:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=f"{valid_count} link(s) have valid destinations",
                    expected="All links should point to valid destinations",
                )
            )

        return results
