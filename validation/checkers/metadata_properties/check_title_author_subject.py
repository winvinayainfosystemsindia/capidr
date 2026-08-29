"""Check: Document properties (Title, Author, Subject) are set.

WCAG 2.2 SC: 2.4.2 Page Titled (A)
"""

from typing import List

from docx import Document

from validation.base_check import BaseCheck
from validation.models import CheckResult


class CheckTitleAuthorSubject(BaseCheck):
    section = "Metadata & Properties"
    checklist_item = "Title, Author, Subject"
    description = "Set Title, Author, and Subject in document properties"
    wcag_criteria = "2.4.2 Page Titled (A)"

    def run(self, doc: Document, doc_path: str) -> List[CheckResult]:
        results = []
        props = doc.core_properties

        # Title
        title = (props.title or "").strip()
        if title:
            results.append(
                self.pass_check(
                    location="File > Properties > Title",
                    actual=f"Title: '{title}'",
                    expected="Document title should be set in properties",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="Document title is not set in properties",
                    location="File > Properties > Title",
                    expected="A descriptive document title in properties",
                    actual="Title is empty",
                )
            )

        # Author
        author = (props.author or "").strip()
        if author:
            results.append(
                self.pass_check(
                    location="File > Properties > Author",
                    actual=f"Author: '{author}'",
                    expected="Author should be set in document properties",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="Document author is not set in properties",
                    location="File > Properties > Author",
                    expected="Author name in document properties",
                    actual="Author is empty",
                )
            )

        # Subject
        subject = (props.subject or "").strip()
        if subject:
            results.append(
                self.pass_check(
                    location="File > Properties > Subject",
                    actual=f"Subject: '{subject}'",
                    expected="Subject should be set in document properties",
                )
            )
        else:
            results.append(
                self.fail_check(
                    reason="Document subject is not set in properties",
                    location="File > Properties > Subject",
                    expected="A descriptive subject in document properties",
                    actual="Subject is empty",
                )
            )

        return results
