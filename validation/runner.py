"""Orchestrator — discovers, runs, and aggregates all validation checkers."""

import importlib
import inspect
import pkgutil
import sys
import traceback
from pathlib import Path
from typing import List, Optional

from docx import Document

from .base_check import BaseCheck
from .models import CheckResult, CheckStatus, ValidationReport
from .report_generator import generate_excel_report


def discover_checkers() -> List[BaseCheck]:
    """Auto-discover all BaseCheck subclasses from the checkers/ package tree.

    Walks every sub-package under ``validation.checkers``, imports each
    module, and collects concrete subclasses of :class:`BaseCheck`.
    """
    checkers: List[BaseCheck] = []
    checkers_package = importlib.import_module("validation.checkers")
    checkers_path = Path(checkers_package.__file__).resolve().parent

    for importer, modname, ispkg in pkgutil.walk_packages(
        path=[str(checkers_path)],
        prefix="validation.checkers.",
    ):
        if ispkg:
            continue  # skip sub-package __init__ files
        try:
            module = importlib.import_module(modname)
        except Exception as exc:
            print(f"[WARNING] Failed to import checker module {modname}: {exc}")
            continue

        for _name, obj in inspect.getmembers(module, inspect.isclass):
            if (
                issubclass(obj, BaseCheck)
                and obj is not BaseCheck
                and not inspect.isabstract(obj)
            ):
                checkers.append(obj())

    # Sort by section name, then checklist item for consistent ordering
    checkers.sort(key=lambda c: (c.section, c.checklist_item))
    return checkers


def run_validation(
    doc_path: str,
    output_path: Optional[str] = None,
    sections: Optional[List[str]] = None,
    verbose: bool = False,
) -> ValidationReport:
    """Run the full validation pipeline.

    Args:
        doc_path: Path to the remediated .docx file.
        output_path: Path for the Excel report. Defaults to
            ``<doc_name>_validation_report.xlsx``.
        sections: If provided, only run checks for these sections.
        verbose: If True, print detailed progress to console.

    Returns:
        The completed :class:`ValidationReport`.
    """
    doc_path_obj = Path(doc_path).resolve()
    if not doc_path_obj.exists():
        raise FileNotFoundError(f"Document not found: {doc_path_obj}")
    if not doc_path_obj.suffix.lower() == ".docx":
        raise ValueError(f"Expected a .docx file, got: {doc_path_obj.suffix}")

    # Determine output path (saved in the same folder as the input document by default)
    if output_path is None:
        output_path = str(
            doc_path_obj.parent / f"{doc_path_obj.stem}_validation_report.xlsx"
        )
    else:
        out_p = Path(output_path)
        if not out_p.is_absolute():
            output_path = str(doc_path_obj.parent / out_p)
        else:
            output_path = str(out_p)

    # Load document once
    if verbose:
        print(f"[INFO] Loading document: {doc_path_obj}")
    doc = Document(str(doc_path_obj))

    # Discover and filter checkers
    all_checkers = discover_checkers()
    if sections:
        sections_lower = {s.lower() for s in sections}
        all_checkers = [c for c in all_checkers if c.section.lower() in sections_lower]

    if verbose:
        print(f"[INFO] Found {len(all_checkers)} checkers to run")

    # Run each checker
    report = ValidationReport(document_path=str(doc_path_obj))
    for checker in all_checkers:
        if verbose:
            print(f"  Running: [{checker.section}] {checker.checklist_item}...")
        try:
            results = checker.run(doc, str(doc_path_obj))
            if results:
                report.add_results(results)
        except Exception as exc:
            # Don't crash — record the error as a FAIL result
            tb = traceback.format_exc()
            if verbose:
                print(f"  [ERROR] {checker.checklist_item}: {exc}")
                print(tb)
            report.add_result(
                CheckResult(
                    section=checker.section,
                    checklist_item=checker.checklist_item,
                    description=checker.description,
                    wcag_criteria=checker.wcag_criteria,
                    status=CheckStatus.FAIL,
                    failure_reason=f"Checker error: {exc}",
                    location="Internal error",
                    expected_value="Check should complete without errors",
                    actual_value=str(exc)[:200],
                )
            )

    # Generate Excel report
    if verbose:
        print(f"\n[INFO] Generating report: {output_path}")
    generate_excel_report(report, output_path)

    # Print summary
    print(f"\n{'=' * 60}")
    print("VALIDATION SUMMARY")
    print(f"{'=' * 60}")
    print(f"Document:  {doc_path_obj.name}")
    print(f"Total Checks: {report.total_checks}")
    print(f"  PASS:    {report.pass_count}")
    print(f"  FAIL:    {report.fail_count}")
    print(f"  N/A:     {report.na_count}")
    print(f"  Pass Rate: {report.pass_rate}% (of applicable checks)")
    print(f"\nReport saved to: {output_path}")
    print(f"{'=' * 60}")

    return report
