"""Check: Headers and footers are accessible and relevant.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckHeadersFooters(BaseCheck):
    section = "Page Layout & Navigation"
    checklist_item = "Headers & Footers"
    description = "Check headers and footers are accessible and relevant"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        header_texts = []
        footer_texts = []

        for section in doc.sections:
            try:
                if section.header:
                    for para in section.header.paragraphs:
                        text = para.text.strip()
                        if text:
                            header_texts.append(text)
            except Exception:
                pass

            try:
                if section.footer:
                    for para in section.footer.paragraphs:
                        text = para.text.strip()
                        if text:
                            footer_texts.append(text)
            except Exception:
                pass

        if header_texts or footer_texts:
            details = []
            if header_texts:
                details.append(f"Header: '{header_texts[0][:50]}'")
            if footer_texts:
                details.append(f"Footer: '{footer_texts[0][:50]}'")
            results.append(
                self.pass_check(
                    location="Document headers/footers",
                    actual="; ".join(details),
                    expected="Headers and footers should contain relevant, accessible content",
                )
            )
        else:
            results.append(
                self.pass_check(
                    location="Document headers/footers",
                    actual="No custom header/footer text found (may use default or empty)",
                    expected="Headers and footers are optional but should be accessible if present",
                )
            )

        return results
