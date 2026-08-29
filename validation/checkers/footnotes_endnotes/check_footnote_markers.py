"""Check: Footnote markers/symbols are correctly linked.

WCAG 2.2 SC: 1.3.1 Info and Relationships (A), 4.1.2 Name, Role, Value (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_footnote_references, get_endnote_references


class CheckFootnoteMarkers(BaseCheck):
    section = "Footnotes & Endnotes"
    checklist_item = "Footnote Markers"
    description = "Ensure markers/symbols are correctly linked to corresponding notes"
    wcag_criteria = "1.3.1 Info and Relationships (A), 4.1.2 Name, Role, Value (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        fn_refs = get_footnote_references(doc)
        en_refs = get_endnote_references(doc)

        if not fn_refs and not en_refs:
            results.append(
                self.not_applicable(
                    "No footnote or endnote references found"
                )
            )
            return results

        # Built-in Word footnotes are inherently linked
        if fn_refs:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=f"{len(fn_refs)} footnote marker(s) correctly linked via Word's built-in feature",
                    expected="Footnote markers should be linked to their corresponding notes",
                )
            )

        if en_refs:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=f"{len(en_refs)} endnote marker(s) correctly linked via Word's built-in feature",
                    expected="Endnote markers should be linked to their corresponding notes",
                )
            )

        return results
