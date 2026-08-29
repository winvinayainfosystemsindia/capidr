"""Data models for the validation framework."""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class CheckStatus(Enum):
    """Status of a single validation check."""

    PASS = "PASS"
    FAIL = "FAIL"
    NOT_APPLICABLE = "N/A"


@dataclass
class CheckResult:
    """Represents the outcome of a single checklist validation item."""

    section: str
    checklist_item: str
    description: str
    wcag_criteria: str
    status: CheckStatus
    failure_reason: str = ""
    location: str = ""
    expected_value: str = ""
    actual_value: str = ""

    def to_row(self) -> list:
        """Convert to a flat list suitable for an Excel row."""
        return [
            self.section,
            self.checklist_item,
            self.description,
            self.wcag_criteria,
            self.status.value,
            self.failure_reason,
            self.location,
            self.expected_value,
            self.actual_value,
        ]


@dataclass
class ValidationReport:
    """Aggregates all check results into a complete validation report."""

    document_path: str
    results: List[CheckResult] = field(default_factory=list)

    @property
    def total_checks(self) -> int:
        return len(self.results)

    @property
    def pass_count(self) -> int:
        return sum(1 for r in self.results if r.status == CheckStatus.PASS)

    @property
    def fail_count(self) -> int:
        return sum(1 for r in self.results if r.status == CheckStatus.FAIL)

    @property
    def na_count(self) -> int:
        return sum(1 for r in self.results if r.status == CheckStatus.NOT_APPLICABLE)

    @property
    def pass_rate(self) -> float:
        """Pass rate calculated over applicable checks only (excludes N/A)."""
        applicable = self.pass_count + self.fail_count
        if applicable == 0:
            return 100.0
        return round((self.pass_count / applicable) * 100, 1)

    def add_result(self, result: CheckResult):
        self.results.append(result)

    def add_results(self, results: List[CheckResult]):
        self.results.extend(results)

    @property
    def headers(self) -> list:
        """Column headers for the Excel detail sheet."""
        return [
            "Section",
            "Checklist Item",
            "Description",
            "WCAG 2.2 Criteria",
            "Status",
            "Failure Reason",
            "Location",
            "Expected Value",
            "Actual Value",
        ]
