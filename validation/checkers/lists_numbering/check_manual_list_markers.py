"""Check: Avoid using manual symbols or numbers for lists.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

import re
from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_nearest_page_marker


class CheckManualListMarkers(BaseCheck):
    section = "Lists & Numbering"
    checklist_item = "Manual List Markers"
    description = "Avoid using manual symbols or numbers for lists"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    # Patterns that indicate sequential manual numbering (3+ consecutive)
    _NUMBERED_PATTERN = re.compile(r"^\d+[.)]\s")

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        sequential_runs = []
        current_run = []

        for idx, para in enumerate(doc.paragraphs):
            style_name = para.style.name if para.style else ""
            text = para.text.strip()

            # Only flag if NOT already using a list style
            if style_name.startswith("List"):
                if current_run and len(current_run) >= 3:
                    sequential_runs.append(current_run[:])
                current_run = []
                continue

            if self._NUMBERED_PATTERN.match(text):
                current_run.append({"index": idx, "text": text[:60]})
            else:
                if current_run and len(current_run) >= 3:
                    sequential_runs.append(current_run[:])
                current_run = []

        # Final check
        if current_run and len(current_run) >= 3:
            sequential_runs.append(current_run[:])

        if not sequential_runs:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="No manual sequential numbering detected",
                    expected="Lists should use Word's list styles, not manual numbering",
                )
            )
        else:
            for run in sequential_runs:
                first = run[0]
                last = run[-1]
                page = get_nearest_page_marker(doc, first["index"])
                results.append(
                    self.fail_check(
                        reason=f"Manual sequential numbering detected ({len(run)} items)",
                        location=(
                            f"Paragraphs {first['index'] + 1}-{last['index'] + 1} "
                            f"(near {page})"
                        ),
                        expected="Use Word's List Number style for numbered lists",
                        actual=f"Manual numbering: \"{first['text']}\" ... \"{last['text']}\"",
                    )
                )

        return results
