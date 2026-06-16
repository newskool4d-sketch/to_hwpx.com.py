from __future__ import annotations

import importlib
import importlib.util
from pathlib import Path
from types import TracebackType
from types import ModuleType, SimpleNamespace
from unittest import SkipTest
from unittest.mock import patch

from parsers import docx as docx_parser
from parsers import pdf as pdf_parser
from parsers import pdf_errors


def _require_module(import_name: str, package_name: str) -> ModuleType:
    if importlib.util.find_spec(import_name) is None:
        raise SkipTest(f"missing package: {package_name}")
    return importlib.import_module(import_name)


def test_docx_parser_preserves_paragraph_and_table_source_order(tmp_path: Path) -> None:
    docx = _require_module("docx", "python-docx")
    document = getattr(docx, "Document")()
    document.add_heading("첫 제목", level=1)
    document.add_paragraph("첫 문단")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "구분"
    table.cell(0, 1).text = "내용"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "첫째"
    document.add_paragraph("마지막 문단")
    path = tmp_path / "ordered.docx"
    document.save(path)

    blocks = docx_parser.parse_docx(str(path))

    assert blocks == [
        {"type": "h", "level": 1, "text": "첫 제목"},
        {"type": "p", "text": "첫 문단"},
        {"type": "table", "header": ["구분", "내용"], "rows": [["A", "첫째"]], "table_source": "docx"},
        {"type": "p", "text": "마지막 문단"},
    ]


def test_docx_parser_preserves_table_source_and_merged_cells(tmp_path: Path) -> None:
    # Given
    docx = _require_module("docx", "python-docx")
    document = getattr(docx, "Document")()
    table = document.add_table(rows=2, cols=3)
    merged_cell = table.cell(0, 0).merge(table.cell(0, 1))
    merged_cell.text = "항목"
    table.cell(0, 2).text = "예산액"
    table.cell(1, 0).text = "강사료"
    table.cell(1, 1).text = "산출내역"
    table.cell(1, 2).text = "400,000원"
    path = tmp_path / "merged.docx"
    document.save(path)

    # When
    blocks = docx_parser.parse_docx(str(path))

    # Then
    assert blocks == [
        {
            "type": "table",
            "header": ["항목", "", "예산액"],
            "rows": [["강사료", "산출내역", "400,000원"]],
            "table_source": "docx",
            "merged_cells": [[0, 0, 1, 2]],
        }
    ]


def test_docx_parser_preserves_vertical_merged_cells(tmp_path: Path) -> None:
    # Given
    docx = _require_module("docx", "python-docx")
    document = getattr(docx, "Document")()
    table = document.add_table(rows=3, cols=2)
    merged_cell = table.cell(0, 0).merge(table.cell(1, 0))
    merged_cell.text = "구분"
    table.cell(0, 1).text = "내용"
    table.cell(1, 1).text = "세부"
    table.cell(2, 0).text = "A"
    table.cell(2, 1).text = "본문"
    path = tmp_path / "vertical-merged.docx"
    document.save(path)

    # When
    blocks = docx_parser.parse_docx(str(path))

    # Then
    assert blocks == [
        {
            "type": "table",
            "header": ["구분", "내용"],
            "rows": [["", "세부"], ["A", "본문"]],
            "table_source": "docx",
            "merged_cells": [[0, 0, 2, 1]],
        }
    ]


def test_pdf_failure_reports_odl_and_all_text_fallback_extractors(tmp_path: Path) -> None:
    pdf_path = tmp_path / "empty.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n%%EOF\n")

    error: RuntimeError | None = None
    with (
        patch.object(
            pdf_parser,
            "extract_pdf_blocks_odl",
            side_effect=pdf_errors.OpendataloaderPdfConversionError("ODL unavailable"),
        ),
        patch.object(pdf_parser, "try_kordoc_pdf_text", return_value=None),
        patch.dict("sys.modules", {"pdfplumber": None, "fitz": None, "pypdf": None}),
    ):
        try:
            pdf_parser.parse_pdf(pdf_path)
        except RuntimeError as exc:
            error = exc

    assert isinstance(error, pdf_errors.PdfTextExtractionError)
    message = str(error)
    assert "opendataloader-pdf: ODL unavailable" in message
    assert "pdfplumber 없음" in message
    assert "PyMuPDF 없음" in message
    assert "pypdf 없음" in message


def test_pdf_parse_does_not_swallow_unexpected_odl_errors(tmp_path: Path) -> None:
    pdf_path = tmp_path / "sample.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n%%EOF\n")
    error: ValueError | None = None

    with (
        patch.object(pdf_parser, "extract_pdf_blocks_odl", side_effect=ValueError("programmer bug")),
        patch.object(pdf_parser, "try_kordoc_pdf_text", return_value="fallback text"),
    ):
        try:
            pdf_parser.parse_pdf(pdf_path)
        except ValueError as exc:
            error = exc

    assert error is not None
    assert "programmer bug" in str(error)


def test_pdf_parse_falls_back_when_odl_convert_api_is_missing(tmp_path: Path) -> None:
    pdf_path = tmp_path / "sample.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n%%EOF\n")

    def fake_import_module(name: str) -> ModuleType | SimpleNamespace:
        if name == "opendataloader_pdf":
            return SimpleNamespace()
        return importlib.import_module(name)

    with (
        patch.object(pdf_parser.importlib, "import_module", side_effect=fake_import_module),
        patch.object(pdf_parser, "try_kordoc_pdf_text", return_value="fallback text"),
    ):
        blocks = pdf_parser.parse_pdf(pdf_path)

    assert blocks == [{"type": "p", "text": "fallback text"}]


def test_pdf_failure_reports_kordoc_command_failure(tmp_path: Path) -> None:
    pdf_path = tmp_path / "sample.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n%%EOF\n")
    failed_process = SimpleNamespace(returncode=2, stdout="", stderr="ocr failed")
    error: RuntimeError | None = None

    with (
        patch.object(
            pdf_parser,
            "extract_pdf_blocks_odl",
            side_effect=pdf_errors.OpendataloaderPdfConversionError("ODL unavailable"),
        ),
        patch.object(pdf_parser, "resolve_kordoc_dir", return_value=tmp_path),
        patch.object(pdf_parser, "kordoc_commands", return_value=[["python", "ocr.py", str(pdf_path)]]),
        patch.object(pdf_parser.subprocess, "run", return_value=failed_process),
        patch.object(
            pdf_parser,
            "extract_pdf_text_fallback",
            side_effect=pdf_errors.PdfTextFallbackError("fallback failed"),
        ),
    ):
        try:
            pdf_parser.parse_pdf(pdf_path)
        except RuntimeError as exc:
            error = exc

    assert isinstance(error, pdf_errors.PdfTextExtractionError)
    message = str(error)
    assert "opendataloader-pdf: ODL unavailable" in message
    assert "kordoc-ai" in message
    assert "exit=2" in message
    assert "ocr failed" in message
    assert "fallback: fallback failed" in message


def test_pdf_text_fallback_does_not_swallow_unexpected_pdfplumber_bug(tmp_path: Path) -> None:
    pdf_path = tmp_path / "buggy.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n%%EOF\n")

    class BuggyPage:
        def extract_text(self) -> str:
            raise AssertionError("programmer bug")

    class BuggyPdf:
        pages = [BuggyPage()]

        def __enter__(self) -> "BuggyPdf":
            return self

        def __exit__(
            self,
            exc_type: type[BaseException] | None,
            exc: BaseException | None,
            traceback: TracebackType | None,
        ) -> bool:
            return False

    def buggy_open(_path: str) -> BuggyPdf:
        return BuggyPdf()

    with patch.dict(
        "sys.modules",
        {"pdfplumber": SimpleNamespace(open=buggy_open), "fitz": None, "pypdf": None},
    ):
        error: AssertionError | None = None
        try:
            pdf_parser.extract_pdf_text_fallback(pdf_path)
        except AssertionError as exc:
            error = exc

    assert error is not None
    assert "programmer bug" in str(error)
