from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
import sys
from typing import TypeVar
from unittest.mock import patch

from parsers.pdf import extract_pdf_text_fallback
import parsers.html as html_parser
import parsers.tabular as tabular_parser


T = TypeVar("T")


def _runtime_error_message(call: Callable[[], T]) -> str:
    try:
        call()
    except RuntimeError as exc:
        return str(exc)
    raise AssertionError("expected RuntimeError")


def test_missing_html_dependency_names_beautifulsoup4() -> None:
    with patch.dict(sys.modules, {"bs4": None}):
        error: html_parser.MissingHtmlDependencyError | None = None
        try:
            html_parser.parse_html("<p>본문</p>")
        except html_parser.MissingHtmlDependencyError as exc:
            error = exc

    assert error is not None
    assert error.package_name == "beautifulsoup4"
    message = str(error)
    assert "beautifulsoup4" in message
    assert "pip install beautifulsoup4" in message


def test_missing_xlsx_dependency_names_openpyxl() -> None:
    with patch.dict(sys.modules, {"openpyxl": None}):
        error: tabular_parser.MissingXlsxDependencyError | None = None
        try:
            tabular_parser.parse_xlsx(Path("sample.xlsx"))
        except tabular_parser.MissingXlsxDependencyError as exc:
            error = exc

    assert error is not None
    assert error.package_name == "openpyxl"
    message = str(error)
    assert "openpyxl" in message
    assert "pip install openpyxl" in message


def test_missing_pdf_fallback_dependencies_name_each_package() -> None:
    blocked_modules = {"pdfplumber": None, "fitz": None, "pypdf": None}

    with patch.dict(sys.modules, blocked_modules):
        message = _runtime_error_message(lambda: extract_pdf_text_fallback(Path("sample.pdf")))

    assert "pdfplumber 없음" in message
    assert "PyMuPDF 없음" in message
    assert "pypdf 없음" in message
