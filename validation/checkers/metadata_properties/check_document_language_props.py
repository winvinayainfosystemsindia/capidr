"""Check: Document language property is set correctly.

WCAG 2.2 SC: 3.1.1 Language of Page (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckDocumentLanguageProps(BaseCheck):
    section = "Metadata & Properties"
    checklist_item = "Document Language Property"
    description = "Set document language in properties"
    wcag_criteria = "3.1.1 Language of Page (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        props = doc.core_properties

        language = (props.language or "").strip()

        if language:
            results.append(
                self.pass_check(
                    location="File > Properties > Language",
                    actual=f"Language property: '{language}'",
                    expected="Document language should be set in core properties",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="Document language property is not set",
                    location="File > Properties > Language",
                    expected="Language should be set (e.g., 'en-US', 'en-GB')",
                    actual="Language property is empty",
                )
            )

        return results
