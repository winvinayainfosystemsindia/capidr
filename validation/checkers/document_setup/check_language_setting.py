"""Check: Document language is set correctly (e.g., English - en-US/en-UK).

WCAG 2.2 SC: 3.1.1 Language of Page (A)
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckLanguageSetting(BaseCheck):
    section = "Document Setup & Structure"
    checklist_item = "Language Setting"
    description = (
        "Verify that the document language is set correctly "
        "(e.g., English - en-US/en-UK)"
    )
    wcag_criteria = "3.1.1 Language of Page (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        # Check document-level language in styles (Normal style)
        body = doc.element.body
        lang_elements = body.findall(f".//{qn('w:lang')}")

        doc_lang = None
        for lang_elem in lang_elements:
            val = lang_elem.get(qn("w:val"))
            if val:
                doc_lang = val
                break

        # Also check the default paragraph style
        try:
            normal_style = doc.styles["Normal"]
            style_elem = normal_style._element
            rpr = style_elem.find(qn("w:rPr"))
            if rpr is not None:
                lang = rpr.find(qn("w:lang"))
                if lang is not None:
                    style_lang = lang.get(qn("w:val"))
                    if style_lang:
                        doc_lang = style_lang
        except Exception:
            pass

        if doc_lang:
            results.append(
                self.pass_check(
                    location="Document Properties > Language / Normal Style",
                    actual=f"Language is set to '{doc_lang}'",
                    expected="Language should be set (e.g., 'en-US', 'en-GB')",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="Document language is not explicitly set",
                    location="Document Properties > Language / Normal Style",
                    expected="Language should be set to appropriate locale (e.g., 'en-US')",
                    actual="No language property found",
                )
            )

        return results
