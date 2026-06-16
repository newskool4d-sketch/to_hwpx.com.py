from __future__ import annotations

from dataclasses import dataclass
from typing import Final
import unicodedata

from blocks import BlockDict
from blocks import BlockValue
from table_grid import int_rows_value
from table_settings import TABLE_TOTAL_WIDTH
from table_settings import calc_col_widths


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
HEADER_FILL_COLOR: Final = 0xE7E7E7

SCHEDULE_HEADER: Final = ("시작", "종료", "분", "내용", "담당")
BUDGET_HEADER_TERMS: Final = ("예산", "예산액", "금액", "단가", "수량", "합계", "원", "amount", "price", "total")
SCHEDULE_HEADER_TERMS: Final = ("시작", "종료", "시간", "시각", "일시", "분", "내용", "담당")
DEFINITION_HEADER_TERMS: Final = ("용어", "정의", "개념", "의미", "뜻", "definition", "term", "meaning")
COMPARISON_HEADER_TERMS: Final = ("현행", "개선", "변경", "전", "후", "비교", "as-is", "to-be", "before", "after")
CONTACT_HEADER_TERMS: Final = ("기관", "부서", "소속", "담당자", "연락처", "전화", "이메일", "contact", "phone", "email")
CHECKLIST_HEADER_TERMS: Final = ("확인", "점검", "체크", "항목", "완료", "여부", "check", "done", "status")
DETAIL_HEADER_TERMS: Final = ("내용", "세부", "내역", "설명", "비고", "추진", "계획")


@dataclass(frozen=True, slots=True)
class TableCellStyle:
    header_align: int = 3
    body_align: int = 1
    header_fill_color: int = HEADER_FILL_COLOR
    border_width: int = 10
    margin_left: int = 170
    margin_right: int = 170
    margin_top: int = 120
    margin_bottom: int = 120


@dataclass(frozen=True, slots=True)
class TableLayout:
    header: list[str]
    rows: list[list[str]]
    table_role: str
    column_widths: list[int]
    style: TableCellStyle
    table_source: str
    merged_cells: list[list[int]]


def _string_list_value(value: BlockValue | None) -> list[str]:
    if not isinstance(value, list):
        return []
    result: list[str] = []
    for item in value:
        if not isinstance(item, str):
            return []
        result.append(item)
    return result


def _string_rows_value(value: BlockValue | None) -> list[list[str]]:
    if not isinstance(value, list):
        return []
    result: list[list[str]] = []
    for row in value:
        if not isinstance(row, list):
            return []
        values: list[str] = []
        for cell in row:
            if not isinstance(cell, str):
                return []
            values.append(cell)
        result.append(values)
    return result


def _int_list_value(value: BlockValue | None) -> list[int]:
    if not isinstance(value, list):
        return []
    result: list[int] = []
    for item in value:
        if type(item) is not int:
            return []
        result.append(item)
    return result


def _col_count(header: list[str], rows: list[list[str]]) -> int:
    all_rows = ([header] if header else []) + rows
    return max((len(row) for row in all_rows), default=0)


def _visual_width(text: str) -> int:
    width = 0
    for char in text:
        if unicodedata.combining(char):
            continue
        width += 2 if unicodedata.east_asian_width(char) in ("F", "W") else 1
    return width


def _scale_widths(widths: list[int], col_count: int, total_width: int) -> list[int]:
    if col_count <= 0 or len(widths) != col_count or any(width <= 0 for width in widths):
        return []
    width_sum = sum(widths)
    if width_sum <= 0:
        return []
    if width_sum == total_width:
        return list(widths)
    scaled = [max(1, int(total_width * width / width_sum)) for width in widths]
    diff = total_width - sum(scaled)
    if diff:
        target = max(range(len(widths)), key=lambda index: widths[index])
        scaled[target] += diff
    return scaled


def _role_profile(table_role: str, col_count: int) -> list[int]:
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
    if len(normalized_header) >= 3 and any(term in joined_header for term in SCHEDULE_HEADER_TERMS):
        return TABLE_ROLE_SCHEDULE
    if len(normalized_header) == 5 and rows and all(":" in row[0] for row in rows if row):
        return TABLE_ROLE_SCHEDULE
    return TABLE_ROLE_BASIC


def _resolved_role(header: list[str], rows: list[list[str]], table_role: str | None) -> str:
    if table_role and table_role.strip():
        return table_role.strip()
    return infer_table_role(header, rows)


def _fit_width_count(widths: list[int], col_count: int, total_width: int) -> list[int]:
    if col_count <= 0:
        return []
    if len(widths) == col_count and sum(widths) == total_width:
        return widths
    if len(widths) != col_count:
        base = [max(1, total_width // col_count) for _ in range(col_count)]
        base[-1] += total_width - sum(base)
        return base
    diff = total_width - sum(widths)
    if diff:
        target = max(range(len(widths)), key=lambda index: widths[index])
        widths[target] += diff
    return widths


def _long_text_target_col(header: list[str], rows: list[list[str]], col_count: int) -> int:
    scores: list[int] = []
    for col_index in range(col_count):
        header_text = header[col_index] if col_index < len(header) else ""
        values = [row[col_index] for row in rows if col_index < len(row)]
        longest = max([_visual_width(header_text), *[_visual_width(value) for value in values]], default=0)
        if any(term in header_text for term in DETAIL_HEADER_TERMS):
            longest += 18
        scores.append(longest)
    if not scores or max(scores) < 48:
        return -1
    return max(range(len(scores)), key=lambda index: scores[index])


def _expand_long_text_width(widths: list[int], header: list[str], rows: list[list[str]], total_width: int) -> list[int]:
    col_count = len(widths)
    target_col = _long_text_target_col(header, rows, col_count)
    if target_col < 0:
        return widths
    target_width = max(widths[target_col], int(total_width * 0.55))
    overflow = target_width - widths[target_col]
    if overflow <= 0:
        return widths
    result = list(widths)
    result[target_col] = target_width
    min_width = max(800, int(total_width * 0.08))
    shrinkable = [index for index in range(col_count) if index != target_col and result[index] > min_width]
    for index in shrinkable:
        if overflow <= 0:
            break
        room = result[index] - min_width
        cut = min(room, max(1, int(overflow / len(shrinkable))))
        result[index] -= cut
        overflow -= cut
    index = 0
    while overflow > 0 and shrinkable:
        col_index = shrinkable[index % len(shrinkable)]
        if result[col_index] > min_width:
            result[col_index] -= 1
            overflow -= 1
        index += 1
        if index > total_width:
            break
    diff = total_width - sum(result)
    result[target_col] += diff
    return result


def table_layout_for(
    header: list[str],
    rows: list[list[str]],
    table_role: str | None = None,
    column_widths: list[int] | None = None,
    table_source: str | None = None,
    merged_cells: list[list[int]] | None = None,
    total_width: int = TABLE_TOTAL_WIDTH,
) -> TableLayout:
    col_count = _col_count(header, rows)
    role = _resolved_role(header, rows, table_role)
    explicit_widths = _scale_widths(column_widths or [], col_count, total_width)
    role_widths = _scale_widths(_role_profile(role, col_count), col_count, total_width)
    calculated_widths = calc_col_widths(header, rows, total=total_width)
    widths = explicit_widths or role_widths or _expand_long_text_width(calculated_widths, header, rows, total_width)
    return TableLayout(
        header=list(header),
        rows=[list(row) for row in rows],
        table_role=role,
        column_widths=_fit_width_count(widths, col_count, total_width),
        style=TableCellStyle(),
        table_source=table_source or "",
        merged_cells=[list(span) for span in (merged_cells or [])],
    )


def table_layout_from_block(block: BlockDict, total_width: int = TABLE_TOTAL_WIDTH) -> TableLayout:
    role_value = block.get("table_role")
    source_value = block.get("table_source")
    return table_layout_for(
        _string_list_value(block.get("header")),
        _string_rows_value(block.get("rows")),
        role_value if isinstance(role_value, str) else None,
        _int_list_value(block.get("column_widths")),
        source_value if isinstance(source_value, str) else None,
        int_rows_value(block.get("merged_cells")),
        total_width=total_width,
    )
