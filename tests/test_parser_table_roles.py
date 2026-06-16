from __future__ import annotations

import importlib
import importlib.util
from pathlib import Path
from types import ModuleType
from unittest import SkipTest

from parsers import pdf as pdf_parser
from parsers.docx import parse_docx
from parsers.html import parse_html
from parsers.tabular import parse_csv_file
from parsers.tabular import parse_xlsx
from table_layout import table_layout_from_block


def _require_module(import_name: str, package_name: str) -> ModuleType:
    if importlib.util.find_spec(import_name) is None:
        raise SkipTest(f"missing package: {package_name}")
    return importlib.import_module(import_name)


def _only_table(blocks):
    tables = [block for block in blocks if block.get("type") == "table"]
    assert len(tables) == 1
    return tables[0]


def test_csv_parser_table_role_is_inferred_from_layout(tmp_path: Path) -> None:
    path = tmp_path / "contacts.csv"
    path.write_text("기관,담당자,연락처\n교육청,홍길동,032-000-0000\n", encoding="utf-8")

    table_block = _only_table(parse_csv_file(path))
    layout = table_layout_from_block(table_block)

    assert table_block["table_source"] == "csv"
    assert layout.table_role == "contacts"


def test_html_parser_table_role_is_inferred_from_layout() -> None:
    _require_module("bs4", "beautifulsoup4")
    html = "<table><tr><th>용어</th><th>정의</th></tr><tr><td>위탁</td><td>외부 기관에 맡김</td></tr></table>"

    table_block = _only_table(parse_html(html))
    layout = table_layout_from_block(table_block)

    assert table_block["table_source"] == "html"
    assert layout.table_role == "definition"


def test_xlsx_parser_table_role_is_inferred_from_layout(tmp_path: Path) -> None:
    openpyxl = _require_module("openpyxl", "openpyxl")
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "체크리스트"
    worksheet.append(["확인", "점검 항목", "비고"])
    worksheet.append(["□", "안전교육 실시", "완료"])
    path = tmp_path / "checklist.xlsx"
    workbook.save(path)
    workbook.close()

    table_block = _only_table(parse_xlsx(path))
    layout = table_layout_from_block(table_block)

    assert table_block["table_source"] == "xlsx"
    assert layout.table_role == "checklist"


def test_docx_parser_table_role_is_inferred_from_layout(tmp_path: Path) -> None:
    docx = _require_module("docx", "python-docx")
    document = docx.Document()
    table = document.add_table(rows=2, cols=3)
    for index, value in enumerate(["구분", "현행", "개선"]):
        table.cell(0, index).text = value
    for index, value in enumerate(["절차", "수기 접수", "온라인 접수"]):
        table.cell(1, index).text = value
    path = tmp_path / "comparison.docx"
    document.save(path)

    table_block = _only_table(parse_docx(str(path)))
    layout = table_layout_from_block(table_block)

    assert table_block["table_source"] == "docx"
    assert layout.table_role == "comparison"


def test_pdf_odl_table_role_is_inferred_from_layout() -> None:
    table_element = {
        "type": "table",
        "rows": [
            {
                "cells": [
                    {"kids": [{"type": "paragraph", "content": "항목"}]},
                    {"kids": [{"type": "paragraph", "content": "산출내역"}]},
                    {"kids": [{"type": "paragraph", "content": "예산액"}]},
                ]
            },
            {
                "cells": [
                    {"kids": [{"type": "paragraph", "content": "강사료"}]},
                    {"kids": [{"type": "paragraph", "content": "2명 x 2시간"}]},
                    {"kids": [{"type": "paragraph", "content": "400,000원"}]},
                ]
            },
        ],
    }

    table_block = _only_table(pdf_parser.odl_element_to_blocks(table_element))
    layout = table_layout_from_block(table_block)

    assert table_block["table_source"] == "pdf"
    assert layout.table_role == "budget"
