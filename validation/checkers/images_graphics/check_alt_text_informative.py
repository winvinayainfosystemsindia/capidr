"""Check: All informative images have descriptive alternative text.

WCAG 2.2 SC: 1.1.1 Non-text Content (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_images, get_nearest_page_marker


class CheckAltTextInformative(BaseCheck):
    section = "Images & Graphics"
    checklist_item = "Alt Text (Informative)"
    description = "Add descriptive alternative text (alt text) for all informative images"
    wcag_criteria = "1.1.1 Non-text Content (A)"

    # Generic alt text patterns that indicate poor alt text
    _BAD_ALT_TEXT = {
        "image", "picture", "photo", "graphic", "figure", "img",
        "screenshot", "chart", "diagram", "icon", "logo",
    }

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        images = get_images(doc)

        if not images:
            results.append(
                self.not_applicable("No images found in the document")
            )
            return results

        for img in images:
            alt_text = img["alt_text"].strip()
            title = img["title"].strip()
            location = (
                f"Paragraph {img['para_index'] + 1}"
                if img["para_index"] >= 0
                else img["para_text"]
            )
            if img["para_index"] >= 0:
                page = get_nearest_page_marker(doc, img["para_index"])
                location = f"{location} (near {page})"

            if not alt_text:
                results.append(
                    self.fail_check(
                        reason="Image has no alt text",
                        location=location,
                        expected="Descriptive alt text describing the image content",
                        actual="Alt text is empty",
                    )
                )
            elif alt_text.lower() in self._BAD_ALT_TEXT:
                results.append(
                    self.fail_check(
                        reason=f"Alt text is too generic: '{alt_text}'",
                        location=location,
                        expected="Meaningful description of the image content",
                        actual=f"Generic alt text: '{alt_text}'",
                    )
                )
            elif len(alt_text) < 5:
                results.append(
                    self.fail_check(
                        reason=f"Alt text is too short: '{alt_text}'",
                        location=location,
                        expected="Descriptive alt text (at least a few words)",
                        actual=f"Alt text: '{alt_text}' ({len(alt_text)} chars)",
                    )
                )
            else:
                results.append(
                    self.pass_check(
                        location=location,
                        actual=f"Alt text: '{alt_text[:80]}'",
                        expected="Descriptive alt text for informative images",
                    )
                )

        return results
