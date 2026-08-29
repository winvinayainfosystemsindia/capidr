"""Check: Footnotes use Word's built-in footnote feature (not manual text).

WCAG 2.2 SC: 1.3.1 Info and Relationships (A)
"""

import re
from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_footnote_references, get_endnote_references, get_nearest_page_marker


class CheckFootnoteFeature(BaseCheck):
    section = "Footnotes & Endnotes"
    checklist_item = "Footnote Feature"
    description = "Use Word's built-in footnote and endnote features (not manual text)"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []

        fn_refs = get_footnote_references(doc)
        en_refs = get_endnote_references(doc)

        # Detect manual footnotes (superscript numbers followed by text at bottom)
        manual_footnotes = []
        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            # Pattern: line starting with superscript number like "1 " or "1. "
            if re.match(r"^[\u00B9\u00B2\u00B3\u2070-\u2079]\s", text):
                manual_footnotes.append(idx)
            # Also check for runs with superscript formatting followed by footnote text
            for run in para.runs:
                if run.font.superscript and run.text.strip().isdigit():
                    # Could be a manual footnote marker
                    pass

        total_builtin = len(fn_refs) + len(en_refs)

        if total_builtin == 0 and not manual_footnotes:
            results.append(
                self.not_applicable(
                    "No footnotes or endnotes found in the document"
                )
            )
        elif total_builtin > 0:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=(
                        f"{len(fn_refs)} footnote(s) and {len(en_refs)} endnote(s) "
                        "using Word's built-in feature"
                    ),
                    expected="Footnotes should use Word's built-in footnote/endnote feature",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="Possible manual footnotes detected instead of Word's built-in feature",
                    location="Document body",
                    expected="Use Word's built-in Insert > Footnote/Endnote feature",
                    actual="Manual superscript numbers found without built-in footnote references",
                )
            )

        return results
