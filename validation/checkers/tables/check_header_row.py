"""Check: Table header rows are marked with Repeat Header Row feature.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_table_info


class CheckHeaderRow(BaseCheck):
    section = "Tables"
    checklist_item = "Header Row"
    description = "Set table headers using the Repeat Header Row feature"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        if not doc.tables:
            results.append(
                self.not_applicable("No tables found in the document")
            )
            return results

        for t_idx, table in enumerate(doc.tables):
            info = get_table_info(table)
            table_label = f"Table {t_idx + 1}"
            header_preview = ", ".join(info["header_texts"][:4]) if info["header_texts"] else "(empty)"

            if info["has_header_row"]:
                results.append(
                    self.pass_check(
                        location=f"{table_label} ({info['row_count']} rows x {info['col_count']} cols)",
                        actual=f"Header row is marked with Repeat Header Row. Headers: {header_preview}",
                        expected="First row should be marked as header with Repeat Header Row",
                    )
                )
            else:
                results.append(
                    self.fail_check(
                        reason="Table header row is not marked with Repeat Header Row",
                        location=f"{table_label} ({info['row_count']} rows x {info['col_count']} cols)",
                        expected="First row should be marked as Repeat Header Row",
                        actual=f"No tblHeader property set. First row: {header_preview}",
                    )
                )

        return results
