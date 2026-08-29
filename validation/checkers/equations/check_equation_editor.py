"""Check: Equations are created using MathType or Word Equation Editor (not images).

WCAG 2.2 SC: 1.1 Non-text Content (A), 4.1.2 Name, Role, Value (A)
"""

from typing import List

from docx import Document
from docx.oxml.ns import qn

from validation.base_check import BaseCheck
from validation.models import CheckResult
from validation.helpers import get_equations, get_images


class CheckEquationEditor(BaseCheck):
    section = "Equations (STEM-specific)"
    checklist_item = "Equation Editor"
    description = "All equations are created using MathType or Word Equation Editor (not as images)"
    wcag_criteria = "1.1 Non-text Content (A), 4.1.2 Name, Role, Value (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        equations = get_equations(doc)
        images = get_images(doc)

        # Check for potential equation images (images with equation-related alt text)
        equation_images = []
        for img in images:
            combined = f"{img['alt_text']} {img['title']} {img['name']}".lower()
            if any(kw in combined for kw in ["equation", "formula", "math", "expr"]):
                equation_images.append(img)

        if not equations and not equation_images:
            results.append(
                self.not_applicable(
                    "No equations found in the document — this check is not applicable"
                )
            )
            return results

        if equations:
            results.append(
                self.pass_check(
                    location="Entire document",
                    actual=f"{len(equations)} native OMML equation(s) found using Word Equation Editor",
                    expected="Equations should use Word Equation Editor (OMML), not images",
                )
            )

        if equation_images:
            for img in equation_images:
                location = (
                    f"Paragraph {img['para_index'] + 1}"
                    if img["para_index"] >= 0
                    else img["para_text"]
                )
                results.append(
                    self.fail_check(
                        reason="Equation appears to be an image instead of a native Word equation",
                        location=location,
                        expected="Use Word Equation Editor (OMML) or MathType",
                        actual=f"Image with equation-related alt text: '{img['alt_text'][:60]}'",
                    )
                )

        return results
