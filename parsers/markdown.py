from __future__ import annotations

import re
from typing import Final, Literal, assert_never

from blocks import BlockDict
from markdown_table_parser import (
    is_markdown_table_separator,
    normalize_parsed_table,
    parse_markdown_table_row,
)

from .common import clean_inline, detect_list_item


BumpisKey = Literal["제목", "소제목", "네모", "원", "바", "별", "당구장", "주석"]
BUMPIS_KEY_BY_RAW: Final[dict[str, BumpisKey]] = {
    "제목": "제목",
    "소제목": "소제목",
    "네모": "네모",
    "원": "원",
    "바": "바",
    "별": "별",
    "당구장": "당구장",
    "주석": "주석",
}


def _split_prefixed_cells(line: str, prefix: str) -> list[str] | None:
    stripped = line.strip()
    marker = f"{prefix}:"
    if not stripped.startswith(marker):
        return None
    cells = [clean_inline(cell.strip()) for cell in stripped[len(marker) :].split(":")]
    return cells if all(cells) else None


def _split_schedule_cells(line: str) -> list[str] | None:
    cells = _split_prefixed_cells(line, "시간계획표")
    if cells is None:
        return None
    if len(cells) == 5 and not all(cell.isdigit() for cell in cells[:4]):
        return cells
    if len(cells) >= 7 and all(cell.isdigit() for cell in cells[:4]):
        start = f"{cells[0]}:{cells[1]}"
        end = f"{cells[2]}:{cells[3]}"
        duration = cells[4]
        content = ":".join(cells[5:-1])
        owner = cells[-1]
        return [start, end, duration, content, owner]
    return None


def _consume_prefixed_table(lines: list[str], start: int, prefix: str) -> tuple[BlockDict, int] | None:
    rows: list[list[str]] = []
    i = start
    while i < len(lines):
        cells = _split_prefixed_cells(lines[i], prefix)
        if cells is None:
            break
        rows.append(cells)
        i += 1
    if not rows:
        return None
    return {"type": "table", "header": rows[0], "rows": rows[1:]}, i


def _consume_schedule_table(lines: list[str], start: int) -> tuple[BlockDict, int] | None:
    rows: list[list[str]] = []
    i = start
    while i < len(lines):
        cells = _split_schedule_cells(lines[i])
        if cells is None:
            break
        rows.append(cells)
        i += 1
    if not rows:
        return None
    return {"type": "table", "header": ["시작", "종료", "분", "내용", "담당"], "rows": rows}, i


def _parse_bumpis_block(line: str) -> BlockDict | None:
    stripped = line.strip()
    if ":" not in stripped:
        return None
    key, raw_value = stripped.split(":", 1)
    value = clean_inline(raw_value.strip())
    if not value:
        return None
    bumpis_key = BUMPIS_KEY_BY_RAW.get(key.strip())
    if bumpis_key is None:
        return None
    match bumpis_key:
        case "제목":
            return {"type": "h", "level": 1, "text": value}
        case "소제목":
            return {"type": "h", "level": 2, "text": value}
        case "네모":
            return {"type": "p", "text": value}
        case "원":
            return {"type": "li", "depth": 0, "marker": "○", "content": value, "text": f"○ {value}"}
        case "바":
            return {"type": "li", "depth": 1, "marker": "─", "content": value, "text": f"─ {value}"}
        case "별":
            return {"type": "li", "depth": 2, "marker": "★", "content": value, "text": f"★ {value}"}
        case "당구장" | "주석":
            return {"type": "bq", "text": f"※ {value}"}
        case unreachable:
            assert_never(unreachable)


def parse_markdown(text: str) -> list[BlockDict]:
    lines = text.splitlines()
    blocks: list[BlockDict] = []
    i = 0
    in_front = False

    while i < len(lines):
        line = lines[i]

        if not line.strip():
            i += 1
            continue

        if line.strip() == "---":
            if i == 0:
                in_front = True
                i += 1
                continue
            if in_front:
                in_front = False
                i += 1
                continue
            blocks.append({"type": "hr"})
            i += 1
            continue

        if in_front:
            i += 1
            continue

        stripped_line = line.strip()

        if stripped_line.startswith("표:"):
            consumed_table = _consume_prefixed_table(lines, i, "표")
            if consumed_table is not None:
                block, i = consumed_table
                blocks.append(block)
                continue

        if stripped_line.startswith("시간계획표:"):
            consumed_schedule = _consume_schedule_table(lines, i)
            if consumed_schedule is not None:
                block, i = consumed_schedule
                blocks.append(block)
                continue

        bumpis_block = _parse_bumpis_block(line)
        if bumpis_block is not None:
            blocks.append(bumpis_block)
            i += 1
            continue

        if re.match(r"^(수신|경유|제목)\s*:", stripped_line):
            colon_idx = stripped_line.index(":")
            key = stripped_line[:colon_idx].strip()
            value = clean_inline(stripped_line[colon_idx + 1 :].strip())
            blocks.append({"type": "official_header", "key": key, "value": value})
            i += 1
            continue

        if re.match(r"^-{3,}\s*$", line) or re.match(r"^\*{3,}\s*$", line):
            blocks.append({"type": "hr"})
            i += 1
            continue

        heading_match = re.match(r"^(#{1,3})\s+(.*)", line)
        if heading_match:
            blocks.append(
                {
                    "type": "h",
                    "level": len(heading_match.group(1)),
                    "text": clean_inline(heading_match.group(2)),
                }
            )
            i += 1
            continue

        if line.strip().startswith("|") and i + 1 < len(lines) and is_markdown_table_separator(lines[i + 1]):
            header = parse_markdown_table_row(line, clean_inline)
            i += 2
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(parse_markdown_table_row(lines[i], clean_inline))
                i += 1
            header, rows = normalize_parsed_table(header, rows)
            blocks.append({"type": "table", "header": header, "rows": rows})
            continue

        item = detect_list_item(line)
        if item:
            blocks.append({"type": "li", **item})
            i += 1
            continue

        if line.strip().startswith(">"):
            blockquote_text = re.sub(r"^>\s*", "", line.strip())
            if blockquote_text:
                blocks.append({"type": "bq", "text": clean_inline(blockquote_text)})
            i += 1
            continue

        if line.strip().startswith("```"):
            i += 1
            code_lines = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1
            for code_line in code_lines:
                if code_line.strip():
                    blocks.append({"type": "code", "text": code_line})
            continue

        paragraph_text = clean_inline(line)
        if paragraph_text:
            blocks.append({"type": "p", "text": paragraph_text})
        i += 1

    return blocks
