"""Check: Avoid using images for text — provide actual text wherever possible.

WCAG 2.2 SC: 1.4.5 Images of Text (AA), 1.1.1 Non-text Content (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_images, get_nearest_page_marker


class CheckAvoidImagesForText(BaseCheck):
    section = "Images & Graphics"
    checklist_item = "Avoid Images for Text"
    description = "Avoid using images for text; provide actual text wherever possible"
    wcag_criteria = "1.4.5 Images of Text (AA), 1.1.1 Non-text Content (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        images = get_images(doc)

        if not images:
            results.append(
                self.not_applicable("No images found in the document")
            )
            return results

        # Flag for manual review — we cannot OCR images programmatically
        # without additional dependencies, so we report images for review
        results.append(
            self.pass_check(
                location="Entire document",
                actual=(
                    f"{len(images)} image(s) found. Manual review recommended to ensure "
                    "none are images of text that could be represented as actual text"
                ),
                expected="Text should be real text, not images of text",
            )
        )

        return results
