from __future__ import annotations

from pathlib import Path

import blocks
from parsers.markdown import parse_markdown
from table_layout import table_layout_from_block


FIXTURE_DIR = Path(__file__).parent / "fixtures"


def test_schedule_fixture_carries_role_and_width_profile() -> None:
    parsed_blocks = parse_markdown((FIXTURE_DIR / "bumpis_syntax.md").read_text(encoding="utf-8"))
    schedule_table = parsed_blocks[-1]

    assert schedule_table["table_role"] == "schedule"
    assert schedule_table["column_widths"] == [14, 14, 9, 49, 14]

    layout = table_layout_from_block(schedule_table)

    assert layout.table_role == "schedule"
    assert len(layout.column_widths) == 5
    assert sum(layout.column_widths) == 14000
    assert layout.column_widths[3] == max(layout.column_widths)


def test_explicit_column_widths_are_normalized_to_total_width() -> None:
    table_block = blocks.table(
        ["구분", "내용", "담당"],
        [["A", "긴 설명", "홍길동"]],
        table_role="basic",
        column_widths=[1, 3, 1],
    )

    layout = table_layout_from_block(table_block, total_width=10000)

    assert layout.table_role == "basic"
    assert layout.column_widths == [2000, 6000, 2000]


def test_budget_table_role_is_inferred_from_amount_headers() -> None:
    table_block = blocks.table(
        ["항목", "수량", "단가", "금액"],
        [["강사료", "2", "100,000원", "200,000원"]],
    )

    layout = table_layout_from_block(table_block)

    assert layout.table_role == "budget"
    assert len(layout.column_widths) == 4
    assert sum(layout.column_widths) == 14000
