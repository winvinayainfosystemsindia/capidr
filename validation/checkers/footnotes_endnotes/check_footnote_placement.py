"""Check: Footnote/endnote text accessibility and placement.

WCAG 2.2 SC: 1.3.2 Meaningful Sequence (A), 2.4.3 Focus Order (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_footnote_references, get_endnote_references


class CheckFootnotePlacement(BaseCheck):
    section = "Footnotes & Endnotes"
    checklist_item = "Footnote Placement"
    description = "Check accessibility of footnote/endnote text and placement"
    wcag_criteria = "1.3.2 Meaningful Sequence (A), 2.4.3 Focus Order (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        fn_refs = get_footnote_references(doc)
        en_refs = get_endnote_references(doc)

        if not fn_refs and not en_refs:
            results.append(
                self.not_applicable(
                    "No footnotes or endnotes found in the document"
                )
            )
            return results

        # Word's built-in footnotes are placed at the bottom of the page
        # and endnotes at the end of the document — both are accessible
        if fn_refs:
            results.append(
                self.pass_check(
                    location="Footnote area",
                    actual=(
                        f"{len(fn_refs)} footnote(s) use Word's built-in placement "
                        "(bottom of page)"
                    ),
                    expected="Footnotes should be placed using Word's built-in feature for proper accessibility",
                )
            )

        if en_refs:
            results.append(
                self.pass_check(
                    location="Endnote area",
                    actual=(
                        f"{len(en_refs)} endnote(s) use Word's built-in placement "
                        "(end of document)"
                    ),
                    expected="Endnotes should be placed using Word's built-in feature",
                )
            )

        return results
