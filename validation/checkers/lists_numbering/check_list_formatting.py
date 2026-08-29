"""Check: Lists use Word's built-in list features (List Bullet / List Number).

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import is_manual_list, get_nearest_page_marker


class CheckListFormatting(BaseCheck):
    section = "Lists & Numbering"
    checklist_item = "List Formatting"
    description = "Ensure bulleted and numbered lists are formatted using Word's list features"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        manual_lists = []

        for idx, para in enumerate(doc.paragraphs):
            marker_type = is_manual_list(para)
            if marker_type:
                page = get_nearest_page_marker(doc, idx)
                manual_lists.append({
                    "index": idx,
                    "text": para.text.strip()[:60],
                    "marker_type": marker_type,
                    "page": page,
                })

        if not manual_lists:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="All lists use Word's built-in list styles",
                    expected="Lists should use List Bullet or List Number styles",
                )
            )
        else:
            # Limit report to first 10, with summary
            shown = manual_lists[:10]
            for ml in shown:
                results.append(
                    self.fail_check(
                        reason=f"Manual {ml['marker_type']} list marker used instead of Word list style",
                        location=f"Paragraph {ml['index'] + 1} (near {ml['page']}), Text: \"{ml['text']}\"",
                        expected="Use Word's built-in List Bullet or List Number style",
                        actual=f"Manual '{ml['marker_type']}' marker typed as text",
                    )
                )
            if len(manual_lists) > 10:
                results.append(
                    self.fail_check(
                        reason=f"{len(manual_lists) - 10} more manual list items found",
                        location="Various locations",
                        expected="Use Word list styles",
                        actual=f"Total: {len(manual_lists)} manual list items",
                    )
                )

        return results
