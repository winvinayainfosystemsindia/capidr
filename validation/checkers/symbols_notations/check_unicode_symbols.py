"""Check: Special symbols use Unicode (not Symbol/Wingdings font or images).

WCAG 2.2 SC: 1.1 Non-text Content (A), 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_nearest_page_marker


class CheckUnicodeSymbols(BaseCheck):
    section = "Symbols, Chemical Notations & Special Characters (STEM-specific)"
    checklist_item = "Unicode Symbols"
    description = "Insert scientific symbols and special characters using Unicode or built-in font characters"
    wcag_criteria = "1.1 Non-text Content (A), 1.3.1 Info and Relationships (A)"

    _SYMBOL_FONTS = {"symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings", "zapf dingbats"}

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        symbol_font_issues = []

        for idx, para in enumerate(doc.paragraphs):
            for run in para.runs:
                if not run.text.strip():
                    continue
                font_name = (run.font.name or "").lower()
                if font_name in self._SYMBOL_FONTS:
                    page = get_nearest_page_marker(doc, idx)
                    symbol_font_issues.append({
                        "index": idx,
                        "font": run.font.name,
                        "text": run.text[:30],
                        "page": page,
                    })

        if not symbol_font_issues:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual="No Symbol/Wingdings fonts detected — Unicode characters used properly",
                    expected="Use Unicode symbols, not Symbol/Wingdings font characters",
                )
            )
        else:
            for issue in symbol_font_issues[:10]:
                results.append(
                    self.fail_check(
                        reason=f"Symbol font '{issue['font']}' used instead of Unicode",
                        location=f"Paragraph {issue['index'] + 1} (near {issue['page']})",
                        expected="Use Unicode character equivalent instead of Symbol/Wingdings font",
                        actual=f"Font: '{issue['font']}', Text: '{issue['text']}'",
                    )
                )

        return results
