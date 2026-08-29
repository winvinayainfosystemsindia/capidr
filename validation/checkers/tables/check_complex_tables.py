"""Check: Complex data tables have descriptions or summaries.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_table_info


class CheckComplexTables(BaseCheck):
    section = "Tables"
    checklist_item = "Complex Tables"
    description = "For complex data tables, use table description or summary for screen reader users"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        if not doc.tables:
            results.append(self.not_applicable("No tables found in the document"))
            return results

        complex_tables = []
        for t_idx, table in enumerate(doc.tables):
            info = get_table_info(table)
            # A table is "complex" if it has many merged cells, many columns, or many rows
            is_complex = (
                len(info["merged_cells"]) > 0
                or info["col_count"] >= 6
                or info["row_count"] >= 15
            )
            if is_complex:
                complex_tables.append((t_idx, info))

        if not complex_tables:
            results.append(
                self.pass_check(
                    location="All tables",
                    actual="No complex tables detected (all tables are simple structure)",
                    expected="Complex tables should have summaries/descriptions",
                )
            )
        else:
            for t_idx, info in complex_tables:
                table_label = f"Table {t_idx + 1}"
                reasons = []
                if len(info["merged_cells"]) > 0:
                    reasons.append(f"{len(info['merged_cells'])} merged cells")
                if info["col_count"] >= 6:
                    reasons.append(f"{info['col_count']} columns")
                if info["row_count"] >= 15:
                    reasons.append(f"{info['row_count']} rows")

                results.append(
                    self.fail_check(
                        reason=f"Complex table detected ({', '.join(reasons)}) — needs description/summary",
                        location=f"{table_label} ({info['row_count']}x{info['col_count']})",
                        expected="Add a table description or summary for screen reader users",
                        actual=f"Complex table with: {', '.join(reasons)}",
                    )
                )

        return results
