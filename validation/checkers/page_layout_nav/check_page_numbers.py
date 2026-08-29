"""Check: Running page numbers are correct.

WCAG 2.2 SC: No direct WCAG mapping.
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckPageNumbers(BaseCheck):
    section = "Page Layout & Navigation"
    checklist_item = "Page Numbers"
    description = "Verify running page numbers are correct"
    wcag_criteria = "No direct WCAG mapping"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        # Check for PAGE field codes in headers/footers
        has_page_numbers = False

        for section in doc.sections:
            for header_footer in [
                section.header,
                section.footer,
                section.even_page_header,
                section.even_page_footer,
                section.first_page_header,
                section.first_page_footer,
            ]:
                if header_footer and header_footer.is_linked_to_previous is False or header_footer:
                    try:
                        for para in header_footer.paragraphs:
                            for fld in para._element.findall(f".//{qn('w:fldChar')}"):
                                has_page_numbers = True
                            for instr in para._element.findall(f".//{qn('w:instrText')}"):
                                if instr.text and "PAGE" in instr.text.upper():
                                    has_page_numbers = True
                    except Exception:
                        pass

        if has_page_numbers:
            results.append(
                self.pass_check(
                    location="Headers/Footers",
                    actual="Page number field codes found in header/footer",
                    expected="Document should have running page numbers",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="No page number field codes found in headers or footers",
                    location="Headers/Footers",
                    expected="Running page numbers in header or footer",
                    actual="No PAGE field codes detected",
                )
            )

        return results
