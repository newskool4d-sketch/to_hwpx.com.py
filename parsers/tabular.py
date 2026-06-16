from __future__ import annotations

import csv
import importlib
from dataclasses import dataclass
from pathlib import Path

from blocks import BlockDict
from table_grid import SourceCell
from table_grid import block_rows_from_grid
from table_grid import expand_spanned_rows

from .common import clean_inline


@dataclass(frozen=True, slots=True)
class MissingXlsxDependencyError(RuntimeError):
    package_name: str
    install_command: str = "pip install openpyxl"

    def __str__(self) -> str:
        return f"XLSX 변환에는 {self.package_name}이 필요함: {self.install_command}"


def parse_csv_file(path: str | Path) -> list[BlockDict]:
    for encoding in ("utf-8-sig", "utf-8", "cp949"):
        try:
            text = Path(path).read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = Path(path).read_text(encoding="utf-8", errors="replace")
    try:
        dialect = csv.Sniffer().sniff(text[:2048])
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(text.splitlines(), dialect))
    rows = [[clean_inline(cell.strip()) for cell in row] for row in rows if any(cell.strip() for cell in row)]
    if not rows:
        return []
    return [{"type": "table", "header": rows[0], "rows": rows[1:], "table_source": "csv"}]


def _cell_text(value) -> str:
    return "" if value is None else clean_inline(str(value))


def _trim_grid(grid: list[list[str]]) -> list[list[str]]:
    while grid and not any(grid[-1]):
        grid.pop()
    max_cols = max((len(row) for row in grid), default=0)
    while max_cols > 0 and all((len(row) <= max_cols - 1 or row[max_cols - 1] == "") for row in grid):
        max_cols -= 1
    return [row[:max_cols] for row in grid]


def _xlsx_table_block(worksheet) -> BlockDict | None:
    top_left_spans: dict[tuple[int, int], tuple[int, int]] = {}
    covered: set[tuple[int, int]] = set()
    for merged_range in worksheet.merged_cells.ranges:
        row_span = merged_range.max_row - merged_range.min_row + 1
        col_span = merged_range.max_col - merged_range.min_col + 1
        top_left_spans[(merged_range.min_row, merged_range.min_col)] = (row_span, col_span)
        for row_index in range(merged_range.min_row, merged_range.max_row + 1):
            for col_index in range(merged_range.min_col, merged_range.max_col + 1):
                if row_index == merged_range.min_row and col_index == merged_range.min_col:
                    continue
                covered.add((row_index, col_index))
    source_rows: list[list[SourceCell]] = []
    for row_index in range(1, worksheet.max_row + 1):
        source_row: list[SourceCell] = []
        for col_index in range(1, worksheet.max_column + 1):
            if (row_index, col_index) in covered:
                continue
            row_span, col_span = top_left_spans.get((row_index, col_index), (1, 1))
            source_row.append(SourceCell(_cell_text(worksheet.cell(row=row_index, column=col_index).value), row_span, col_span))
        if source_row:
            source_rows.append(source_row)
    grid, merged_cells = expand_spanned_rows(source_rows)
    trimmed_grid = _trim_grid(grid)
    header, rows = block_rows_from_grid(trimmed_grid)
    if not header:
        return None
    block: BlockDict = {
        "type": "table",
        "header": header,
        "rows": rows,
        "table_source": "xlsx",
        "worksheet_title": worksheet.title,
    }
    if merged_cells:
        block["merged_cells"] = merged_cells
    return block


def parse_xlsx(path: str | Path) -> list[BlockDict]:
    try:
        openpyxl = importlib.import_module("openpyxl")
    except ImportError as exc:
        raise MissingXlsxDependencyError(package_name="openpyxl") from exc
    load_workbook = getattr(openpyxl, "load_workbook")
    workbook = load_workbook(path, data_only=True)
    try:
        blocks: list[BlockDict] = []
        for worksheet in workbook.worksheets:
            table_block = _xlsx_table_block(worksheet)
            if table_block is None:
                continue
            blocks.append({"type": "h", "level": 2, "text": worksheet.title})
            blocks.append(table_block)
        return blocks
    finally:
        workbook.close()
