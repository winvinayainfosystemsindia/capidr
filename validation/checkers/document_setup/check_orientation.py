"""Check: Document orientation is correct (portrait/landscape as appropriate).

WCAG 2.2 SC: 1.3.4 Orientation (AA)
"""

from typing import List

from docx import Document
from docx.enum.section import WD_ORIENT

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckOrientation(BaseCheck):
    section = "Document Setup & Structure"
    checklist_item = "Orientation"
    description = "Ensure document orientation is correct (portrait/landscape as appropriate)"
    wcag_criteria = "1.3.4 Orientation (AA)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        for idx, section in enumerate(doc.sections):
            sec_label = f"Section {idx + 1}" if len(doc.sections) > 1 else "Document"
            orientation = section.orientation

            if orientation == WD_ORIENT.PORTRAIT:
                orient_str = "Portrait"
            elif orientation == WD_ORIENT.LANDSCAPE:
                orient_str = "Landscape"
            else:
                # Determine from dimensions
                width = section.page_width or 0
                height = section.page_height or 0
                orient_str = "Portrait" if height >= width else "Landscape"

            # WCAG 1.3.4: Content should not be restricted to a single orientation.
            # We report the orientation; a mixed-orientation doc is fine.
            results.append(
                self.pass_check(
                    location=f"{sec_label} > Page Setup > Orientation",
                    actual=f"Orientation: {orient_str}",
                    expected="Content should not be locked to a single orientation",
                )
            )

        return results
