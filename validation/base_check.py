"""Abstract base class for all validation checkers."""

from abc import ABC, abstractmethod
from typing import List

from docx import Document

from .models import CheckResult, CheckStatus


class BaseCheck(ABC):
    """Abstract base class that all individual checkers must inherit from.

    Each checker defines metadata (section, checklist_item, description,
    wcag_criteria) and implements the `run()` method that performs the
    actual validation against a loaded Word document.
    """

    # Subclasses MUST set these class-level attributes
    section: str = ""
    checklist_item: str = ""
    description: str = ""
    wcag_criteria: str = ""

    @abstractmethod
    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        """Execute the validation check against the given document.

        Args:
            doc: A loaded python-docx Document object.
            doc_path: Absolute path to the .docx file on disk.

        Returns:
            A list of CheckResult objects. Most checks return a single
            result, but some may return multiple (e.g., one per table).
        """
        ...

    # ------------------------------------------------------------------
    # Convenience factory methods for creating results
    # ------------------------------------------------------------------

    def pass_check(
        self, location: str = "", actual: str = "", expected: str = ""
    ) -> CheckResult:
        """Create a PASS result."""
        return CheckResult(
            section=self.section,
            checklist_item=self.checklist_item,
            description=self.description,
            wcag_criteria=self.wcag_criteria,
            status=CheckStatus.PASS,
            location=location,
            expected_value=expected,
            actual_value=actual,
        )

    def fail_check(
        self,
        reason: str,
        location: str = "",
        expected: str = "",
        actual: str = "",
    ) -> CheckResult:
        """Create a FAIL result."""
        return CheckResult(
            section=self.section,
            checklist_item=self.checklist_item,
            description=self.description,
            wcag_criteria=self.wcag_criteria,
            status=CheckStatus.FAIL,
            failure_reason=reason,
            location=location,
            expected_value=expected,
            actual_value=actual,
        )

    def not_applicable(self, reason: str = "") -> CheckResult:
        """Create a NOT_APPLICABLE result."""
        return CheckResult(
            section=self.section,
            checklist_item=self.checklist_item,
            description=self.description,
            wcag_criteria=self.wcag_criteria,
            status=CheckStatus.NOT_APPLICABLE,
            failure_reason=reason if reason else "Not applicable to this document",
            location="",
            expected_value="",
            actual_value="",
        )
