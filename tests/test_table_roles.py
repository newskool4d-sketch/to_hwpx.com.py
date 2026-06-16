from __future__ import annotations

from table_roles import TABLE_ROLE_CHECKLIST
from table_roles import TABLE_ROLE_COMPARISON
from table_roles import TABLE_ROLE_CONTACTS
from table_roles import TABLE_ROLE_DEFINITION
from table_roles import TABLE_ROLE_SCHEDULE
from table_roles import infer_table_role
from table_roles import role_column_widths


def test_infer_table_role_recognizes_public_document_patterns() -> None:
    # Given
    cases = [
        (["시작", "종료", "분", "내용", "담당"], [], TABLE_ROLE_SCHEDULE),
        (["용어", "정의"], [["위탁", "사무를 외부 기관에 맡기는 방식"]], TABLE_ROLE_DEFINITION),
        (["구분", "현행", "개선"], [["절차", "수기", "온라인"]], TABLE_ROLE_COMPARISON),
        (["기관", "담당자", "연락처"], [["교육청", "홍길동", "032-000-0000"]], TABLE_ROLE_CONTACTS),
        (["확인", "점검 항목", "비고"], [["□", "안전교육 실시", "완료"]], TABLE_ROLE_CHECKLIST),
    ]

    for header, rows, expected_role in cases:
        # When
        role = infer_table_role(header, rows)

        # Then
        assert role == expected_role


def test_role_column_widths_returns_known_profiles_by_column_count() -> None:
    assert role_column_widths(TABLE_ROLE_DEFINITION, 2) == [30, 70]
    assert role_column_widths(TABLE_ROLE_COMPARISON, 3) == [24, 38, 38]
    assert role_column_widths(TABLE_ROLE_CONTACTS, 3) == [42, 24, 34]
    assert role_column_widths(TABLE_ROLE_CHECKLIST, 3) == [14, 66, 20]
    assert role_column_widths("unknown", 3) == []
