"""Check: Tables have no merged/split cells that affect logical reading order.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A), 1.3.2 Meaningful Sequence (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_table_info


class CheckMergedCells(BaseCheck):
    section = "Tables"
    checklist_item = "Merged/Split Cells"
    description = "Ensure tables have no merged/split cells that affect logical reading order"
    wcag_criteria = "1.3.1 Info and Relationships (A), 1.3.2 Meaningful Sequence (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        if not doc.tables:
            results.append(self.not_applicable("No tables found in the document"))
            return results

        for t_idx, table in enumerate(doc.tables):
            info = get_table_info(table)
            table_label = f"Table {t_idx + 1}"
            merged = info["merged_cells"]

            if not merged:
                results.append(
                    self.pass_check(
                        location=f"{table_label} ({info['row_count']}x{info['col_count']})",
                        actual="No merged or split cells detected",
                        expected="Tables should not have merged cells",
                    )
                )
            else:
                merge_details = []
                for mc in merged[:5]:
                    merge_details.append(
                        f"Row {mc['row'] + 1}, Col {mc['col'] + 1} ({mc['type']} merge)"
                    )
                results.append(
                    self.fail_check(
                        reason=f"{len(merged)} merged cell(s) found that may affect reading order",
                        location=f"{table_label}: {'; '.join(merge_details)}",
                        expected="No merged cells — use simple table structure",
                        actual=f"{len(merged)} merged cell(s) detected",
                    )
                )

        return results
