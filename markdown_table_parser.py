import re
from collections.abc import Callable
from typing import Final


SEPARATOR_PATTERN: Final = re.compile(r'^[ \t]*:?-+:?[ \t]*$')


def split_markdown_table_cells(line: str) -> list[str]:
    stripped = line.strip()
    if stripped.startswith('|'):
        stripped = stripped[1:]
    if stripped.endswith('|'):
        stripped = stripped[:-1]

    cells: list[str] = []
    current: list[str] = []
    escaped = False
    for char in stripped:
        if escaped:
            if char == '|':
                current.append('|')
            else:
                current.append('\\')
                current.append(char)
            escaped = False
            continue
        if char == '\\':
            escaped = True
            continue
        if char == '|':
            cells.append(''.join(current))
            current = []
            continue
        current.append(char)
    if escaped:
        current.append('\\')
    cells.append(''.join(current))
    return cells


def is_markdown_table_separator(line: str) -> bool:
    if len(line) > 500:
        return False
    cells = split_markdown_table_cells(line)
    return bool(cells) and all(SEPARATOR_PATTERN.match(cell) for cell in cells)


def parse_markdown_table_row(line: str, clean_inline: Callable[[str], str]) -> list[str]:
    return [clean_inline(cell.strip()) for cell in split_markdown_table_cells(line)]


def normalize_parsed_table(header: list[str], rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    all_rows = [header, *rows]
    col_count = max((len(row) for row in all_rows), default=0)
    if col_count == 0:
        return header, rows
    normalized = [row + [''] * (col_count - len(row)) for row in all_rows]
    return normalized[0], normalized[1:]
