from __future__ import annotations

import importlib
import importlib.util
from pathlib import Path
from types import ModuleType
from unittest import SkipTest

import blocks
from parsers import pdf as pdf_parser
from parsers.html import parse_html
from parsers.tabular import parse_xlsx
from table_layout import table_layout_from_block


def _require_module(import_name: str, package_name: str) -> ModuleType:
    if importlib.util.find_spec(import_name) is None:
        raise SkipTest(f"missing package: {package_name}")
    return importlib.import_module(import_name)


def test_layout_expands_long_text_column_before_generic_columns() -> None:
    # Given
    table_block = blocks.table(
        ["구분", "세부 내용", "담당"],
        [
            [
                "A",
                "학교 현장의 수요와 지역 여건을 반영한 긴 설명 문장입니다. " * 4,
                "홍길동",
            ]
        ],
    )

    # When
    layout = table_layout_from_block(table_block, total_width=14000)

    # Then
    assert layout.column_widths[1] >= 7600
    assert layout.column_widths[1] > layout.column_widths[2]


def test_budget_preset_keeps_amount_column_compact_and_detail_column_wide() -> None:
    # Given
    table_block = blocks.table(
        ["항목", "산출내역", "예산액"],
        [["강사료", "2명 x 2시간 x 100,000원", "400,000원"]],
    )

    # When
    layout = table_layout_from_block(table_block, total_width=10000)

    # Then
    assert layout.table_role == "budget"
    assert layout.column_widths == [2400, 5000, 2600]


def test_public_document_table_roles_are_inferred_from_headers() -> None:
    # Given
    cases = [
        (["용어", "정의"], [["위탁", "사무를 외부 기관에 맡기는 방식"]], "definition", 1),
        (["구분", "현행", "개선"], [["절차", "수기 접수", "온라인 접수"]], "comparison", 1),
        (["기관", "담당자", "연락처"], [["교육청", "홍길동", "032-000-0000"]], "contacts", 0),
        (["확인", "점검 항목", "비고"], [["□", "안전교육 실시", "완료"]], "checklist", 1),
    ]

    for header, rows, expected_role, wide_col in cases:
        # When
        layout = table_layout_from_block(blocks.table(header, rows), total_width=10000)

        # Then
        assert layout.table_role == expected_role
        assert sum(layout.column_widths) == 10000
        assert layout.column_widths[wide_col] == max(layout.column_widths)


def test_html_parser_preserves_source_and_cell_spans() -> None:
    # Given
    _require_module("bs4", "beautifulsoup4")
    html = (
        "<table>"
        '<tr><th rowspan="2">구분</th><th colspan="2">내용</th></tr>'
        "<tr><th>세부</th><th>담당</th></tr>"
        "<tr><td>A</td><td>긴 내용</td><td>홍길동</td></tr>"
        "</table>"
    )

    # When
    parsed_blocks = parse_html(html)

    # Then
    assert parsed_blocks == [
        {
            "type": "table",
            "header": ["구분", "내용", ""],
            "rows": [["", "세부", "담당"], ["A", "긴 내용", "홍길동"]],
            "table_source": "html",
            "merged_cells": [[0, 0, 2, 1], [0, 1, 1, 2]],
        }
    ]


def test_xlsx_parser_preserves_source_sheet_and_merged_cells(tmp_path: Path) -> None:
    # Given
    openpyxl = _require_module("openpyxl", "openpyxl")
    workbook = getattr(openpyxl, "Workbook")()
    worksheet = workbook.active
    worksheet.title = "예산"
    worksheet.merge_cells("A1:B1")
    worksheet["A1"] = "항목"
    worksheet["C1"] = "예산액"
    worksheet.append(["강사료", "산출내역", "400,000원"])
    path = tmp_path / "merged.xlsx"
    workbook.save(path)
    workbook.close()

    # When
    parsed_blocks = parse_xlsx(path)

    # Then
    assert parsed_blocks == [
        {"type": "h", "level": 2, "text": "예산"},
        {
            "type": "table",
            "header": ["항목", "", "예산액"],
            "rows": [["강사료", "산출내역", "400,000원"]],
            "table_source": "xlsx",
            "worksheet_title": "예산",
            "merged_cells": [[0, 0, 1, 2]],
        },
    ]


def test_pdf_odl_parser_preserves_source_and_cell_spans() -> None:
    # Given
    table_element = {
        "type": "table",
        "rows": [
            {
                "cells": [
                    {"rowspan": 2, "kids": [{"type": "paragraph", "content": "구분"}]},
                    {"colspan": 2, "kids": [{"type": "paragraph", "content": "내용"}]},
                ]
            },
            {"cells": [{"kids": [{"type": "paragraph", "content": "세부"}]}, {"kids": [{"type": "paragraph", "content": "담당"}]}]},
            {
                "cells": [
                    {"kids": [{"type": "paragraph", "content": "A"}]},
                    {"kids": [{"type": "paragraph", "content": "긴 내용"}]},
                    {"kids": [{"type": "paragraph", "content": "홍길동"}]},
                ]
            },
        ],
    }

    # When
    parsed_blocks = pdf_parser.odl_element_to_blocks(table_element)

    # Then
    assert parsed_blocks == [
        {
            "type": "table",
            "header": ["구분", "내용", ""],
            "rows": [["", "세부", "담당"], ["A", "긴 내용", "홍길동"]],
            "table_source": "pdf",
            "merged_cells": [[0, 0, 2, 1], [0, 1, 1, 2]],
        }
    ]
