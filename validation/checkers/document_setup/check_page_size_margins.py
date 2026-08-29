"""Check: Page size and margins meet publisher/organization requirements.

WCAG 2.2 SC: No direct WCAG mapping — publisher/organizational standard.
"""

from typing import List

from docx import Document
from docx.shared import Inches

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckPageSizeMargins(BaseCheck):
    section = "Document Setup & Structure"
    checklist_item = "Page Size & Margins"
    description = (
        "Confirm the page size and margins meet publisher/organization requirements"
    )
    wcag_criteria = "No direct WCAG mapping"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        for idx, section in enumerate(doc.sections):
            sec_label = f"Section {idx + 1}" if len(doc.sections) > 1 else "Document"

            # Page size
            width_in = round(section.page_width / 914400, 2)  # EMU to inches
            height_in = round(section.page_height / 914400, 2)

            # Margins
            left_in = round(section.left_margin / 914400, 2) if section.left_margin else 0
            right_in = round(section.right_margin / 914400, 2) if section.right_margin else 0
            top_in = round(section.top_margin / 914400, 2) if section.top_margin else 0
            bottom_in = round(section.bottom_margin / 914400, 2) if section.bottom_margin else 0

            size_info = f"Page: {width_in}\" x {height_in}\""
            margin_info = (
                f"Margins: L={left_in}\", R={right_in}\", "
                f"T={top_in}\", B={bottom_in}\""
            )

            # Standard US Letter = 8.5 x 11, A4 ≈ 8.27 x 11.69
            is_letter = abs(width_in - 8.5) < 0.1 and abs(height_in - 11.0) < 0.1
            is_a4 = abs(width_in - 8.27) < 0.15 and abs(height_in - 11.69) < 0.15

            if is_letter or is_a4:
                page_type = "US Letter" if is_letter else "A4"
                results.append(
                    self.pass_check(
                        location=f"{sec_label} > Page Setup",
                        actual=f"{size_info} ({page_type}), {margin_info}",
                        expected="Standard page size (US Letter or A4) with appropriate margins",
                    )
                )
            else:
                results.append(
                    self.fail_check(
                        reason=f"Non-standard page size detected: {width_in}\" x {height_in}\"",
                        location=f"{sec_label} > Page Setup",
                        expected="Standard page size: US Letter (8.5\" x 11\") or A4 (8.27\" x 11.69\")",
                        actual=f"{size_info}, {margin_info}",
                    )
                )

        return results
