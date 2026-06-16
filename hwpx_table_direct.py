from __future__ import annotations

from blocks import BlockValue
from table_grid import int_rows_value
from table_layout import table_layout_for
from table_settings import calc_row_heights


TABLE_BODY_WIDTH = 42520
TABLE_MIN_RENDER_HEIGHT = 2200


def string_list_value(value: BlockValue | None) -> list[str]:
    if not isinstance(value, list):
        return []
    result: list[str] = []
    for item in value:
        if not isinstance(item, str):
            return []
        result.append(item)
    return result


def string_rows_value(value: BlockValue | None) -> list[list[str]]:
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


def int_list_value(value: BlockValue | None) -> list[int]:
    if not isinstance(value, list):
        return []
    result: list[int] = []
    for item in value:
        if type(item) is not int:
            return []
        result.append(item)
    return result


def _normalize_rows(header: list[str], rows: list[list[str]]) -> tuple[list[list[str]], int]:
    all_rows = ([header] if header else []) + (rows or [])
    col_count = max((len(row) for row in all_rows), default=0)
    normalized: list[list[str]] = []
    for row in all_rows:
        values = [str(cell or "").strip() for cell in row]
        normalized.append(values + [""] * (col_count - len(values)))
    return normalized, col_count


def _span_map(merged_cells: list[list[int]]) -> dict[tuple[int, int], tuple[int, int]]:
    result: dict[tuple[int, int], tuple[int, int]] = {}
    for row, col, row_span, col_span in merged_cells:
        result[(row, col)] = (row_span, col_span)
    return result


def _make_table_cell(text, col_idx, row_idx, width, height, is_header, helpers, row_span=1, col_span=1):
    cell_para_id = helpers.next_id()
    border_fill = "13" if is_header else "4"
    char_pr = "18" if is_header else "38"
    para_pr = "21" if is_header else "4"
    header_flag = "1" if is_header else "0"
    return (
        f'<hp:tc name="" header="{header_flag}" hasMargin="1" protect="0" editable="0" '
        f'dirty="1" borderFillIDRef="{border_fill}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" '
        f'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" '
        f'hasTextRef="0" hasNumRef="0">'
        f'<hp:p paraPrIDRef="{para_pr}" styleIDRef="0" pageBreak="0" columnBreak="0" '
        f'merged="0" id="{cell_para_id}">'
        f'<hp:run charPrIDRef="{char_pr}"><hp:t>{helpers.xml_escape(text)}</hp:t></hp:run>'
        f"</hp:p></hp:subList>"
        f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
        f'<hp:cellSpan colSpan="{col_span}" rowSpan="{row_span}"/>'
        f'<hp:cellSz width="{width}" height="{height}"/>'
        f'<hp:cellMargin left="170" right="170" top="120" bottom="120"/></hp:tc>'
    )


def make_table_xml(header, rows, helpers, table_role=None, column_widths=None, merged_cells=None):
    normalized_header = string_list_value(header)
    normalized_rows = string_rows_value(rows)
    layout = table_layout_for(
        normalized_header or [],
        normalized_rows or [],
        table_role if isinstance(table_role, str) else None,
        int_list_value(column_widths),
        merged_cells=int_rows_value(merged_cells),
        total_width=TABLE_BODY_WIDTH,
    )
    normalized, col_count = _normalize_rows(layout.header, layout.rows)
    if not normalized or col_count == 0:
        return ""
    widths = layout.column_widths
    heights = calc_row_heights(layout.header, layout.rows, widths[:col_count])
    if len(heights) < len(normalized):
        heights.extend([TABLE_MIN_RENDER_HEIGHT] * (len(normalized) - len(heights)))
    heights = [max(TABLE_MIN_RENDER_HEIGHT, height) for height in heights]
    spans = _span_map(layout.merged_cells)
    p_id = helpers.next_id()
    tbl_id = helpers.next_id()
    table_rows: list[str] = []
    for row_idx, row in enumerate(normalized):
        cells: list[str] = []
        is_header = bool(layout.header) and row_idx == 0
        for col_idx in range(col_count):
            row_span, col_span = spans.get((row_idx, col_idx), (1, 1))
            cell_width = sum(widths[col_idx : min(col_count, col_idx + col_span)])
            cells.append(_make_table_cell(row[col_idx], col_idx, row_idx, cell_width, heights[row_idx], is_header, helpers, row_span, col_span))
        table_rows.append(f'<hp:tr>{"".join(cells)}</hp:tr>')
    total_height = sum(heights)
    return (
        f'<hp:p id="{p_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">'
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
        f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" '
        f'rowCnt="{len(normalized)}" colCnt="{col_count}" cellSpacing="0" borderFillIDRef="4" noAdjust="0">'
        f'<hp:sz width="{TABLE_BODY_WIDTH}" widthRelTo="ABSOLUTE" height="{total_height}" '
        f'heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
        f'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
        f'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'{"".join(table_rows)}</hp:tbl></hp:run></hp:p>'
    )
