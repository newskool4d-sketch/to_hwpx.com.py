from __future__ import annotations

import blocks


def test_block_helpers_match_current_dict_contract() -> None:
    assert blocks.heading(1, "제목") == {"type": "h", "level": 1, "text": "제목"}
    assert blocks.paragraph("본문") == {"type": "p", "text": "본문"}
    assert blocks.list_item("1. 항목", depth=1, marker="1.", content="항목") == {
        "type": "li",
        "depth": 1,
        "text": "1. 항목",
        "marker": "1.",
        "content": "항목",
    }
    assert blocks.table(["구분", "내용"], [["A", "첫째"]]) == {
        "type": "table",
        "header": ["구분", "내용"],
        "rows": [["A", "첫째"]],
    }
    assert blocks.blockquote("인용") == {"type": "bq", "text": "인용"}
    assert blocks.code("print('hi')") == {"type": "code", "text": "print('hi')"}
    assert blocks.horizontal_rule() == {"type": "hr"}
    assert blocks.official_header("수신", "홍길동") == {
        "type": "official_header",
        "key": "수신",
        "value": "홍길동",
    }


def test_table_helper_copies_mutable_inputs() -> None:
    header = ["구분", "내용"]
    rows = [["A", "첫째"]]

    block = blocks.table(header, rows)
    header.append("추가")
    rows[0].append("변경")

    assert block == {"type": "table", "header": ["구분", "내용"], "rows": [["A", "첫째"]]}


def test_invalid_block_type_is_rejected() -> None:
    error: blocks.UnsupportedBlockTypeError | None = None
    try:
        blocks.ensure_supported_block_type("unknown")
    except blocks.UnsupportedBlockTypeError as exc:
        error = exc

    assert error is not None
    assert error.block_type == "unknown"
    assert "지원하지 않는 block type" in str(error)
