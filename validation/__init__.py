"""WCAG 2.2 Remediated Document Validation Framework.

Takes a remediated .docx file as input, runs all accessibility checklist
items, and produces a detailed Excel validation report.
"""

from .runner import run_validation

__all__ = ["run_validation"]
