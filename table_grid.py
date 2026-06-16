from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class SourceCell:
    text: str
    row_span: int = 1
    col_span: int = 1


def _normalized_span(value: int) -> int:
    return max(1, value)


def _ensure_grid_cell(grid: list[list[str]], row: int, col: int) -> None:
    while len(grid) <= row:
        grid.append([])
    while len(grid[row]) <= col:
        grid[row].append("")


def expand_spanned_rows(rows: list[list[SourceCell]]) -> tuple[list[list[str]], list[list[int]]]:
    grid: list[list[str]] = []
    covered: set[tuple[int, int]] = set()
    merged_cells: list[list[int]] = []
    for row_index, source_row in enumerate(rows):
        _ensure_grid_cell(grid, row_index, 0)
        col_index = 0
        for source_cell in source_row:
            while (row_index, col_index) in covered:
                _ensure_grid_cell(grid, row_index, col_index)
                col_index += 1
            row_span = _normalized_span(source_cell.row_span)
            col_span = _normalized_span(source_cell.col_span)
            _ensure_grid_cell(grid, row_index, col_index)
            grid[row_index][col_index] = source_cell.text
            if row_span > 1 or col_span > 1:
                merged_cells.append([row_index, col_index, row_span, col_span])
            for target_row in range(row_index, row_index + row_span):
                for target_col in range(col_index, col_index + col_span):
                    _ensure_grid_cell(grid, target_row, target_col)
                    if target_row == row_index and target_col == col_index:
                        continue
                    grid[target_row][target_col] = ""
                    covered.add((target_row, target_col))
            col_index += col_span
    max_cols = max((len(row) for row in grid), default=0)
    return [row + [""] * (max_cols - len(row)) for row in grid], merged_cells


def block_rows_from_grid(grid: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    if not grid:
        return [], []
    return list(grid[0]), [list(row) for row in grid[1:]]


def int_rows_value(value: str | int | list[str] | list[int] | list[list[str]] | list[list[int]] | None) -> list[list[int]]:
    if not isinstance(value, list):
        return []
    result: list[list[int]] = []
    for row in value:
        if not isinstance(row, list):
            return []
        values: list[int] = []
        for cell in row:
            if type(cell) is not int:
                return []
            values.append(cell)
        if len(values) == 4 and values[2] > 0 and values[3] > 0:
            result.append(values)
    return result
