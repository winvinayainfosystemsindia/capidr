"""Check: Decorative images are marked as decorative (empty alt text or marked).

WCAG 2.2 SC: 1.1.1 Non-text Content (A)
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_images


class CheckDecorativeImages(BaseCheck):
    section = "Images & Graphics"
    checklist_item = "Decorative Images"
    description = "Decorative images marked as decorative in alt text or left blank"
    wcag_criteria = "1.1.1 Non-text Content (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        images = get_images(doc)

        if not images:
            results.append(
                self.not_applicable("No images found in the document")
            )
            return results

        # Check for images that might be decorative
        # (lines, separators, small icons, borders)
        decorative_count = 0
        informative_count = 0

        for img in images:
            alt_text = img["alt_text"].strip()
            name = img["name"].strip().lower()

            # Heuristic: images with names suggesting decorative purpose
            decorative_names = ["line", "separator", "border", "spacer", "divider", "decoration"]
            is_likely_decorative = any(d in name for d in decorative_names)

            if is_likely_decorative:
                decorative_count += 1
                if alt_text == "" or alt_text.lower() in ("decorative", ""):
                    # Correctly marked as decorative
                    pass
                else:
                    results.append(
                        self.fail_check(
                            reason="Likely decorative image has non-empty alt text",
                            location=f"Image: '{name}'",
                            expected="Decorative images should have empty alt text or be marked decorative",
                            actual=f"Alt text: '{alt_text}'",
                        )
                    )
            else:
                informative_count += 1

        if decorative_count == 0:
            results.append(
                self.pass_check(
                    location="All images",
                    actual=f"All {informative_count} image(s) appear to be informative",
                    expected="Decorative images should be marked appropriately",
                )
            )

        return results
