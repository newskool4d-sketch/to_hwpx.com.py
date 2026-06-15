from __future__ import annotations

import importlib
import importlib.util
import json
from pathlib import Path
from types import ModuleType
from unittest import SkipTest
from unittest.mock import patch

import to_hwpx_com
from blocks import BlockDict
from parsers import pdf as pdf_parser
from parsers import pdf_errors


FIXTURE_DIR = Path(__file__).parent / "fixtures"
EXPECTED_BLOCKS_PATH = FIXTURE_DIR / "parser_expected.json"


def _expected_blocks(name: str) -> list[BlockDict]:
    loaded = json.loads(EXPECTED_BLOCKS_PATH.read_text(encoding="utf-8"))
    if not isinstance(loaded, dict):
        raise AssertionError("parser_expected.json must contain an object")
    return loaded[name]


def _require_module(import_name: str, package_name: str) -> ModuleType:
    if importlib.util.find_spec(import_name) is None:
        raise SkipTest(f"missing package: {package_name}")
    return importlib.import_module(import_name)


def test_markdown_fixture_blocks() -> None:
    blocks = to_hwpx_com.detect_and_parse(FIXTURE_DIR / "parser_sample.md")

    assert blocks == _expected_blocks("markdown")


def test_txt_fixture_blocks() -> None:
    blocks = to_hwpx_com.detect_and_parse(FIXTURE_DIR / "parser_sample.txt")

    assert blocks == _expected_blocks("txt")


def test_html_fixture_blocks() -> None:
    _require_module("bs4", "beautifulsoup4")

    blocks = to_hwpx_com.detect_and_parse(FIXTURE_DIR / "parser_sample.html")

    assert blocks == _expected_blocks("html")


def test_csv_fixture_blocks() -> None:
    blocks = to_hwpx_com.detect_and_parse(FIXTURE_DIR / "parser_sample.csv")

    assert blocks == _expected_blocks("csv")


def test_empty_csv_fixture_returns_empty_blocks() -> None:
    blocks = to_hwpx_com.detect_and_parse(FIXTURE_DIR / "parser_empty.csv")

    assert blocks == _expected_blocks("empty_csv")


def test_xlsx_fixture_blocks(tmp_path: Path) -> None:
    openpyxl = _require_module("openpyxl", "openpyxl")
    workbook = getattr(openpyxl, "Workbook")()
    worksheet = workbook.active
    worksheet.title = "자료"
    worksheet.append(["구분", "내용"])
    worksheet.append(["A", "첫째"])
    worksheet.append(["B", "둘째"])
    path = tmp_path / "parser_sample.xlsx"
    workbook.save(path)
    workbook.close()

    blocks = to_hwpx_com.detect_and_parse(path)

    assert blocks == _expected_blocks("xlsx")


def test_docx_fixture_blocks(tmp_path: Path) -> None:
    docx = _require_module("docx", "python-docx")
    document = getattr(docx, "Document")()
    document.add_heading("DOCX 제목", level=1)
    document.add_paragraph("수신: 홍길동")
    document.add_paragraph("본문")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "구분"
    table.cell(0, 1).text = "내용"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "첫째"
    path = tmp_path / "parser_sample.docx"
    document.save(path)

    blocks = to_hwpx_com.detect_and_parse(path)

    assert blocks == _expected_blocks("docx")


def test_pdf_text_fallback_fixture_blocks(tmp_path: Path) -> None:
    fitz = _require_module("fitz", "PyMuPDF")
    pdf = getattr(fitz, "open")()
    page = pdf.new_page()
    page.insert_text((72, 72), "PDF Title\nPDF body")
    path = tmp_path / "parser_sample.pdf"
    pdf.save(str(path))
    pdf.close()

    with (
        patch.object(
            pdf_parser,
            "extract_pdf_blocks_odl",
            side_effect=pdf_errors.OpendataloaderPdfConversionError("forced fallback fixture"),
        ),
        patch.object(pdf_parser, "try_kordoc_pdf_text", return_value=None),
    ):
        blocks = to_hwpx_com.detect_and_parse(path)

    assert blocks == _expected_blocks("pdf")
