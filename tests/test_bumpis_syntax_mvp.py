from __future__ import annotations

from pathlib import Path

from parsers.markdown import parse_markdown


FIXTURE_DIR = Path(__file__).parent / "fixtures"


def test_bumpis_syntax_fixture_parses_to_existing_block_types() -> None:
    blocks = parse_markdown((FIXTURE_DIR / "bumpis_syntax.md").read_text(encoding="utf-8"))

    assert blocks == [
        {"type": "h", "level": 1, "text": "영상회의 개최 계획"},
        {"type": "h", "level": 2, "text": "회의 개요"},
        {"type": "p", "text": "추진 배경"},
        {"type": "li", "depth": 0, "marker": "○", "content": "참석 대상 안내", "text": "○ 참석 대상 안내"},
        {"type": "li", "depth": 1, "marker": "─", "content": "세부 내용", "text": "─ 세부 내용"},
        {"type": "li", "depth": 2, "marker": "★", "content": "참고 사항", "text": "★ 참고 사항"},
        {"type": "bq", "text": "※ 일정은 변동될 수 있음"},
        {"type": "bq", "text": "※ 내부 검토용"},
        {"type": "table", "header": ["구분", "내용"], "rows": [["A", "첫째"], ["B", "둘째"]]},
        {
            "type": "table",
            "header": ["시작", "종료", "분", "내용", "담당"],
            "rows": [["15:00", "15:05", "5’", "인사 말씀", "국장"], ["15:05", "15:25", "20’", "보고", "과장"]],
        },
    ]


def test_malformed_bumpis_schedule_line_falls_back_to_paragraph() -> None:
    blocks = parse_markdown("시간계획표:10:00:10:10:활동만 있음\n")

    assert blocks == [{"type": "p", "text": "시간계획표:10:00:10:10:활동만 있음"}]
