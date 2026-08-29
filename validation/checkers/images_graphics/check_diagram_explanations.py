"""Check: Diagrams and scientific figures have accessible explanations.

WCAG 2.2 SC: 1.1.1 Non-text Content (A)
"""

import re
from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_images, get_nearest_page_marker


class CheckDiagramExplanations(BaseCheck):
    section = "Images & Graphics"
    checklist_item = "Diagram Explanations"
    description = "Provide accessible explanations for diagrams and scientific figures"
    wcag_criteria = "1.1.1 Non-text Content (A)"

    _DIAGRAM_KEYWORDS = re.compile(
        r"(diagram|figure|illustration|schematic|flow\s*chart|circuit|structure|anatomy|model)",
        re.IGNORECASE,
    )

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        images = get_images(doc)

        if not images:
            results.append(self.not_applicable("No images found in the document"))
            return results

        diagram_images = []
        for img in images:
            combined = f"{img['alt_text']} {img['title']} {img['name']}"
            if self._DIAGRAM_KEYWORDS.search(combined):
                diagram_images.append(img)

        if not diagram_images:
            results.append(
                self.not_applicable(
                    "No diagrams or scientific figures detected"
                )
            )
        else:
            for diag in diagram_images:
                alt_text = diag["alt_text"].strip()
                location = (
                    f"Paragraph {diag['para_index'] + 1}"
                    if diag["para_index"] >= 0
                    else diag["para_text"]
                )
                if diag["para_index"] >= 0:
                    page = get_nearest_page_marker(doc, diag["para_index"])
                    location = f"{location} (near {page})"

                if alt_text and len(alt_text) >= 15:
                    results.append(
                        self.pass_check(
                            location=location,
                            actual=f"Diagram has explanation: '{alt_text[:80]}'",
                            expected="Diagrams should have accessible text explanations",
                        )
                    )
                else:
                    results.append(
                        self.fail_check(
                            reason="Diagram/figure lacks an accessible text explanation",
                            location=location,
                            expected="Provide an accessible text explanation for the diagram",
                            actual=f"Alt text: '{alt_text}' (insufficient for a diagram)",
                        )
                    )

        return results
