"""Check: Nested lists use correct hierarchy and indentation.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_nearest_page_marker


class CheckNestedLists(BaseCheck):
    section = "Lists & Numbering"
    checklist_item = "Nested Lists"
    description = "Check that nested lists use correct hierarchy and indentation"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        list_items = []
        nesting_issues = []

        for idx, para in enumerate(doc.paragraphs):
            style_name = para.style.name if para.style else ""
            if not style_name.startswith("List"):
                continue

            # Get indentation level (ilvl)
            ilvl = 0
            ppr = para._element.find(qn("w:pPr"))
            if ppr is not None:
                num_pr = ppr.find(qn("w:numPr"))
                if num_pr is not None:
                    ilvl_elem = num_pr.find(qn("w:ilvl"))
                    if ilvl_elem is not None:
                        try:
                            ilvl = int(ilvl_elem.get(qn("w:val"), "0"))
                        except ValueError:
                            pass

            list_items.append({"index": idx, "level": ilvl, "text": para.text.strip()[:50]})

        if not list_items:
            results.append(
                self.not_applicable("No list items found in the document")
            )
            return results

        # Check for proper nesting (no jump more than 1 level)
        prev_level = 0
        for i, item in enumerate(list_items):
            if item["level"] > prev_level + 1:
                page = get_nearest_page_marker(doc, item["index"])
                nesting_issues.append({
                    "index": item["index"],
                    "expected_max": prev_level + 1,
                    "actual": item["level"],
                    "text": item["text"],
                    "page": page,
                })
            prev_level = item["level"]

        if not nesting_issues:
            max_depth = max(item["level"] for item in list_items)
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=f"All {len(list_items)} list items have proper nesting (max depth: {max_depth})",
                    expected="List nesting should not skip indentation levels",
                )
            )
        else:
            for issue in nesting_issues:
                results.append(
                    self.fail_check(
                        reason=f"List nesting skips levels: jumped to level {issue['actual']}",
                        location=(
                            f"Paragraph {issue['index'] + 1} (near {issue['page']}), "
                            f"Text: \"{issue['text']}\""
                        ),
                        expected=f"Maximum level {issue['expected_max']} (one deeper than parent)",
                        actual=f"Level {issue['actual']}",
                    )
                )

        return results
