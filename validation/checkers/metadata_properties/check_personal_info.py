"""Check: Remove personal or sensitive information from document properties.

WCAG 2.2 SC: No direct WCAG mapping (privacy best practice).
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckPersonalInfo(BaseCheck):
    section = "Metadata & Properties"
    checklist_item = "Personal Information"
    description = "Remove personal or sensitive information from properties"
    wcag_criteria = "No direct WCAG mapping"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        props = doc.core_properties

        # Check for potentially sensitive fields
        sensitive_fields = {}
        if props.comments:
            sensitive_fields["Comments"] = props.comments[:80]
        if props.last_modified_by:
            sensitive_fields["Last Modified By"] = props.last_modified_by

        # Check custom properties in XML
        try:
            custom_props_part = doc.part.package.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
            )
            if custom_props_part:
                sensitive_fields["Custom Properties"] = "Custom properties exist"
        except Exception:
            pass

        if not sensitive_fields:
            results.append(
                self.pass_check(
                    location="Document Properties",
                    actual="No personal or sensitive information detected in properties",
                    expected="Document should not contain personal/sensitive metadata",
                )
            )
        else:
            for field, value in sensitive_fields.items():
                results.append(
                    self.pass_check(
                        location=f"Document Properties > {field}",
                        actual=f"{field}: '{value}'",
                        expected="Review for personal/sensitive information",
                    )
                )

        return results
