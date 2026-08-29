"""Check: Non-standard symbols have explanations or nomenclature.

WCAG 2.2 SC: 3.1.3 Unusual Words (AAA)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckNonstandardSymbols(BaseCheck):
    section = "Symbols, Chemical Notations & Special Characters (STEM-specific)"
    checklist_item = "Non-standard Symbols"
    description = "Provide explanations for non-standard symbols and nomenclature"
    wcag_criteria = "3.1.3 Unusual Words (AAA)"

    # Unicode ranges for special/technical symbols
    _SPECIAL_RANGES = [
        (0x2200, 0x22FF),  # Mathematical Operators
        (0x2300, 0x23FF),  # Miscellaneous Technical
        (0x2600, 0x26FF),  # Miscellaneous Symbols
        (0x2700, 0x27BF),  # Dingbats
        (0x2100, 0x214F),  # Letterlike Symbols
    ]

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        special_symbols = set()

        for para in doc.paragraphs:
            for char in para.text:
                code = ord(char)
                for start, end in self._SPECIAL_RANGES:
                    if start <= code <= end:
                        special_symbols.add(char)

        if not special_symbols:
            results.append(
                self.not_applicable(
                    "No non-standard technical symbols detected"
                )
            )
        else:
            symbols_str = " ".join(sorted(special_symbols)[:20])
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=(
                        f"{len(special_symbols)} unique special symbol(s) found: {symbols_str}. "
                        "Manual review recommended to ensure explanations are provided "
                        "where needed."
                    ),
                    expected="Non-standard symbols should have explanations or be defined",
                )
            )

        return results
