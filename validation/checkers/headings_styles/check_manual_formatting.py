"""Check: Document uses styles only (no manual formatting for structure).

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document
from docx.shared import Pt

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_nearest_page_marker, is_heading


class CheckManualFormatting(BaseCheck):
    section = "Headings & Styles"
    checklist_item = "Manual Formatting"
    description = "Remove manual formatting — use styles only for structure"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        manual_format_issues = []

        for idx, para in enumerate(doc.paragraphs):
            if is_heading(para):
                continue  # headings are fine
            text = para.text.strip()
            if not text:
                continue

            for run in para.runs:
                run_text = run.text.strip()
                if not run_text:
                    continue

                # Detect direct font size overrides (not inherited from style)
                if run.font.size and run.font.size != Pt(12):
                    font_pt = round(run.font.size / 12700, 1)
                    page = get_nearest_page_marker(doc, idx)
                    manual_format_issues.append({
                        "index": idx,
                        "page": page,
                        "issue": f"Direct font size: {font_pt}pt",
                        "text": run_text[:40],
                    })
                    break  # one issue per paragraph is enough

        if not manual_format_issues:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="No significant manual formatting detected in body text",
                    expected="Use styles for structure; avoid direct formatting overrides",
                )
            )
        else:
            # Group and limit to first 10 to keep report manageable
            shown = manual_format_issues[:10]
            for issue in shown:
                results.append(
                    self.fail_check(
                        reason=f"Manual formatting used: {issue['issue']}",
                        location=(
                            f"Paragraph {issue['index'] + 1} (near {issue['page']}), "
                            f"Text: \"{issue['text']}\""
                        ),
                        expected="Use styles for formatting, not direct overrides",
                        actual=issue["issue"],
                    )
                )
            if len(manual_format_issues) > 10:
                results.append(
                    self.fail_check(
                        reason=f"{len(manual_format_issues) - 10} more manual formatting issues found",
                        location="Various locations throughout the document",
                        expected="Use styles for formatting",
                        actual=f"Total: {len(manual_format_issues)} paragraphs with manual formatting",
                    )
                )

        return results
