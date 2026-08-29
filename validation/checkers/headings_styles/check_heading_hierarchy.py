"""Check: Heading hierarchy is correct (no skipped levels).

WCAG 2.2 SC: 1.3.1 Info and Relationships (A), 2.4.10 Section Headings (AAA)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_headings


class CheckHeadingHierarchy(BaseCheck):
    section = "Headings & Styles"
    checklist_item = "Heading Hierarchy"
    description = (
        "Check hierarchical order of headings is correct (no skipped heading levels)"
    )
    wcag_criteria = "1.3.1 Info and Relationships (A), 2.4.10 Section Headings (AAA)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        headings = get_headings(doc)

        # Filter out H6 page markers
        content_headings = [(idx, lvl, txt) for idx, lvl, txt in headings if lvl != 6]

        if not content_headings:
            results.append(
                self.fail_check(
                    reason="No headings found in the document",
                    location="Entire document",
                    expected="Document should have at least one heading (H1)",
                    actual="No Heading styles applied to any paragraph",
                )
            )
            return results

        # Check: should have exactly one H1
        h1_count = sum(1 for _, lvl, _ in content_headings if lvl == 1)
        if h1_count == 0:
            results.append(
                self.fail_check(
                    reason="No Heading 1 found in the document",
                    location="Document structure",
                    expected="Exactly one Heading 1 (document title)",
                    actual=f"Found {h1_count} Heading 1(s)",
                )
            )
        elif h1_count > 1:
            h1_locations = [
                f"Paragraph {idx + 1}: \"{txt[:40]}\""
                for idx, lvl, txt in content_headings
                if lvl == 1
            ]
            results.append(
                self.fail_check(
                    reason=f"Multiple Heading 1s found ({h1_count})",
                    location="; ".join(h1_locations),
                    expected="Exactly one Heading 1 (document title)",
                    actual=f"{h1_count} Heading 1(s) found",
                )
            )
        else:
            results.append(
                self.pass_check(
                    location="Document structure",
                    actual="Exactly one Heading 1 found",
                    expected="Exactly one Heading 1 (document title)",
                )
            )

        # Check: no skipped levels
        skip_issues = []
        prev_level = 0
        for idx, lvl, txt in content_headings:
            if prev_level > 0 and lvl > prev_level + 1:
                skip_issues.append({
                    "index": idx,
                    "text": txt[:50],
                    "expected_max": prev_level + 1,
                    "actual": lvl,
                    "prev": prev_level,
                })
            prev_level = lvl

        if not skip_issues:
            results.append(
                self.pass_check(
                    location="Document structure",
                    actual="No skipped heading levels detected",
                    expected="Heading levels should not skip (e.g., H1 → H3 without H2)",
                )
            )
        else:
            for issue in skip_issues:
                results.append(
                    self.fail_check(
                        reason=(
                            f"Skipped heading level: jumped from H{issue['prev']} to H{issue['actual']}"
                        ),
                        location=f"Paragraph {issue['index'] + 1}: \"{issue['text']}\"",
                        expected=f"Heading {issue['expected_max']} or lower (no skipping)",
                        actual=f"Heading {issue['actual']}",
                    )
                )

        return results
