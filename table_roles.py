from __future__ import annotations

from typing import Final


TABLE_ROLE_BASIC: Final = "basic"
TABLE_ROLE_SCHEDULE: Final = "schedule"
TABLE_ROLE_BUDGET: Final = "budget"
TABLE_ROLE_DEFINITION: Final = "definition"
TABLE_ROLE_COMPARISON: Final = "comparison"
TABLE_ROLE_CONTACTS: Final = "contacts"
TABLE_ROLE_CHECKLIST: Final = "checklist"

SCHEDULE_COLUMN_WIDTHS: Final = (14, 14, 9, 49, 14)
BUDGET_COLUMN_WIDTHS_BY_COUNT: Final = {3: (24, 50, 26), 4: (22, 42, 20, 16)}
ROLE_COLUMN_WIDTHS_BY_COUNT: Final = {
    (TABLE_ROLE_DEFINITION, 2): (30, 70),
    (TABLE_ROLE_DEFINITION, 3): (22, 50, 28),
    (TABLE_ROLE_COMPARISON, 3): (24, 38, 38),
    (TABLE_ROLE_COMPARISON, 4): (18, 28, 28, 26),
    (TABLE_ROLE_CONTACTS, 3): (42, 24, 34),
    (TABLE_ROLE_CONTACTS, 4): (34, 22, 24, 20),
    (TABLE_ROLE_CHECKLIST, 3): (14, 66, 20),
    (TABLE_ROLE_CHECKLIST, 4): (12, 58, 16, 14),
}

SCHEDULE_HEADER: Final = ("시작", "종료", "분", "내용", "담당")
BUDGET_HEADER_TERMS: Final = ("예산", "예산액", "금액", "단가", "수량", "합계", "원", "amount", "price", "total")
SCHEDULE_HEADER_TERMS: Final = ("시작", "종료", "시간", "시각", "일시", "분")
DEFINITION_HEADER_TERMS: Final = ("용어", "정의", "개념", "의미", "뜻", "definition", "term", "meaning")
COMPARISON_HEADER_TERMS: Final = ("현행", "개선", "변경", "전", "후", "비교", "as-is", "to-be", "before", "after")
CONTACT_HEADER_TERMS: Final = ("기관", "부서", "소속", "담당자", "연락처", "전화", "이메일", "contact", "phone", "email")
CHECKLIST_HEADER_TERMS: Final = ("확인", "점검", "체크", "항목", "완료", "여부", "check", "done", "status")
DETAIL_HEADER_TERMS: Final = ("내용", "세부", "내역", "설명", "비고", "추진", "계획")


def _has_schedule_header_term(header: tuple[str, ...]) -> bool:
    for cell in header:
        for term in SCHEDULE_HEADER_TERMS:
            if len(term) == 1 and cell == term:
                return True
            if len(term) > 1 and term in cell:
                return True
    return False


def role_column_widths(table_role: str, col_count: int) -> list[int]:
    if table_role == TABLE_ROLE_SCHEDULE and col_count == len(SCHEDULE_COLUMN_WIDTHS):
        return list(SCHEDULE_COLUMN_WIDTHS)
    if table_role == TABLE_ROLE_BUDGET:
        return list(BUDGET_COLUMN_WIDTHS_BY_COUNT.get(col_count, ()))
    profile = ROLE_COLUMN_WIDTHS_BY_COUNT.get((table_role, col_count), ())
    if profile:
        return list(profile)
    return []


def infer_table_role(header: list[str], rows: list[list[str]]) -> str:
    normalized_header = tuple(cell.strip() for cell in header)
    if normalized_header == SCHEDULE_HEADER:
        return TABLE_ROLE_SCHEDULE
    joined_header = " ".join(normalized_header).lower()
    if any(term in joined_header for term in BUDGET_HEADER_TERMS):
        return TABLE_ROLE_BUDGET
    if sum(term in joined_header for term in DEFINITION_HEADER_TERMS) >= 2:
        return TABLE_ROLE_DEFINITION
    if sum(term in joined_header for term in COMPARISON_HEADER_TERMS) >= 2:
        return TABLE_ROLE_COMPARISON
    if sum(term in joined_header for term in CONTACT_HEADER_TERMS) >= 2:
        return TABLE_ROLE_CONTACTS
    if sum(term in joined_header for term in CHECKLIST_HEADER_TERMS) >= 2:
        return TABLE_ROLE_CHECKLIST
    if len(normalized_header) >= 3 and _has_schedule_header_term(normalized_header):
        return TABLE_ROLE_SCHEDULE
    if len(normalized_header) == 5 and rows and all(":" in row[0] for row in rows if row):
        return TABLE_ROLE_SCHEDULE
    return TABLE_ROLE_BASIC
