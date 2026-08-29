"""Check: Table of Contents uses Word's automatic TOC feature.

WCAG 2.2 SC: 2.4.5 Multiple Ways (AA), 1.3.1 Info and Relationships (A)
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckTableOfContents(BaseCheck):
    section = "Page Layout & Navigation"
    checklist_item = "Table of Contents"
    description = "Generate an automatic Table of Contents and update fields"
    wcag_criteria = "2.4.5 Multiple Ways (AA), 1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        # Look for TOC field codes
        has_toc = False
        for para in doc.paragraphs:
            for instr in para._element.findall(f".//{qn('w:instrText')}"):
                if instr.text and "TOC" in instr.text.upper():
                    has_toc = True
                    break
            # Also check for TOC styles
            style_name = para.style.name if para.style else ""
            if style_name.startswith("TOC") or "toc" in style_name.lower():
                has_toc = True

        # Check for structured document tags (SDT) containing TOC
        for sdt in doc.element.body.findall(f".//{qn('w:sdt')}"):
            for doc_part in sdt.findall(f".//{qn('w:sdtPr')}"):
                for alias in doc_part.findall(qn("w:alias")):
                    val = alias.get(qn("w:val"), "")
                    if "TOC" in val.upper() or "table of contents" in val.lower():
                        has_toc = True

        if has_toc:
            results.append(
                self.pass_check(
                    location="Document structure",
                    actual="Automatic Table of Contents (TOC) field detected",
                    expected="Document should have an automatic TOC with hyperlinks",
                )
            )
        else:
            # Check if there are enough headings to warrant a TOC
            heading_count = sum(
                1 for para in doc.paragraphs
                if para.style and para.style.name.startswith("Heading")
                and para.style.name != "Heading 6"
            )
            if heading_count >= 3:
                results.append(
                    self.fail_check(
                        reason="No automatic Table of Contents found",
                        location="Document structure",
                        expected="Generate an automatic TOC using Word's References > Table of Contents",
                        actual=f"No TOC found, but document has {heading_count} headings",
                    )
                )
            else:
                results.append(
                    self.not_applicable(
                        f"Document has only {heading_count} heading(s) — TOC may not be necessary"
                    )
                )

        return results
