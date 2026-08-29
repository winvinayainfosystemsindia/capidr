"""Check: Hyperlinks have meaningful display text (not raw URLs).

WCAG 2.2 SC: 2.4.4 Link Purpose (In Context) (A), 2.4.9 Link Purpose (Link Only) (AAA)
"""

import re
from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_hyperlinks, get_nearest_page_marker


class CheckLinkDisplayText(BaseCheck):
    section = "Hyperlinks & Cross References"
    checklist_item = "Link Display Text"
    description = "Ensure all hyperlinks have meaningful display text (not raw URLs)"
    wcag_criteria = "2.4.4 Link Purpose (In Context) (A), 2.4.9 Link Purpose (Link Only) (AAA)"

    _BAD_LINK_TEXTS = {
        "click here", "here", "link", "more", "read more",
        "this link", "click", "go", "see more",
    }
    _URL_PATTERN = re.compile(r"^https?://")

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        links = get_hyperlinks(doc)

        if not links:
            results.append(
                self.not_applicable("No hyperlinks found in the document")
            )
            return results

        for link in links:
            display = link["display_text"].strip()
            url = link["url"]
            location = link["location"]
            if link["para_index"] >= 0:
                page = get_nearest_page_marker(doc, link["para_index"])
                location = f"{location} (near {page})"

            if not display:
                results.append(
                    self.fail_check(
                        reason="Hyperlink has no display text",
                        location=location,
                        expected="Meaningful, descriptive link text",
                        actual=f"Empty display text, URL: '{url[:60]}'",
                    )
                )
            elif self._URL_PATTERN.match(display):
                results.append(
                    self.fail_check(
                        reason="Hyperlink displays a raw URL instead of descriptive text",
                        location=location,
                        expected="Meaningful display text describing the link destination",
                        actual=f"Raw URL as display text: '{display[:60]}'",
                    )
                )
            elif display.lower() in self._BAD_LINK_TEXTS:
                results.append(
                    self.fail_check(
                        reason=f"Generic link text: '{display}'",
                        location=location,
                        expected="Descriptive text that explains where the link goes",
                        actual=f"Generic text: '{display}'",
                    )
                )
            else:
                results.append(
                    self.pass_check(
                        location=location,
                        actual=f"Link text: '{display[:40]}' → {url[:40]}",
                        expected="Meaningful display text for hyperlinks",
                    )
                )

        return results
