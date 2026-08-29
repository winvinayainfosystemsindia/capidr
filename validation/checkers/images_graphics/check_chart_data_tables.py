"""Check: Charts and graphs have detailed image descriptions or data tables.

WCAG 2.2 SC: 1.1.1 Non-text Content (A)
"""

import re
from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_images, get_nearest_page_marker


class CheckChartDataTables(BaseCheck):
    section = "Images & Graphics"
    checklist_item = "Chart Data Tables"
    description = "Ensure charts and graphs have detailed image descriptions or data tables"
    wcag_criteria = "1.1.1 Non-text Content (A)"

    _CHART_KEYWORDS = re.compile(
        r"(chart|graph|plot|histogram|pie\s+chart|bar\s+chart|scatter|trend|data\s+vis)",
        re.IGNORECASE,
    )

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        images = get_images(doc)

        if not images:
            results.append(self.not_applicable("No images found in the document"))
            return results

        chart_images = []
        for img in images:
            alt_text = img["alt_text"]
            title = img["title"]
            name = img["name"]
            combined = f"{alt_text} {title} {name}"
            if self._CHART_KEYWORDS.search(combined):
                chart_images.append(img)

        if not chart_images:
            results.append(
                self.not_applicable(
                    "No charts or graphs detected in the document"
                )
            )
        else:
            for chart in chart_images:
                alt_text = chart["alt_text"].strip()
                location = (
                    f"Paragraph {chart['para_index'] + 1}"
                    if chart["para_index"] >= 0
                    else chart["para_text"]
                )
                if chart["para_index"] >= 0:
                    page = get_nearest_page_marker(doc, chart["para_index"])
                    location = f"{location} (near {page})"

                if alt_text and len(alt_text) >= 20:
                    results.append(
                        self.pass_check(
                            location=location,
                            actual=f"Chart has description: '{alt_text[:80]}'",
                            expected="Charts should have detailed descriptions or accompanying data tables",
                        )
                    )
                else:
                    results.append(
                        self.fail_check(
                            reason="Chart/graph lacks a detailed description or data table",
                            location=location,
                            expected="Provide a detailed text description or accessible data table for the chart",
                            actual=f"Alt text: '{alt_text}' (insufficient for a chart)",
                        )
                    )

        return results
