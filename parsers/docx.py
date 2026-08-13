from __future__ import annotations

import importlib
import re
from typing import Literal, assert_never

from blocks import BlockDict
from table_grid import MAX_TABLE_SPAN, block_rows_from_grid


DocxBlockTag = Literal["paragraph", "table"]


def iter_block_items(doc):
    ns_module = importlib.import_module("docx.oxml.ns")
    table_module = importlib.import_module("docx.table")
    paragraph_module = importlib.import_module("docx.text.paragraph")
    qn = getattr(ns_module, "qn")
    docx_table = getattr(table_module, "Table")
    docx_paragraph = getattr(paragraph_module, "Paragraph")
    block_tag_by_raw: dict[str, DocxBlockTag] = {qn("w:p"): "paragraph", qn("w:tbl"): "table"}

    body = doc.element.body
    for child in body.iterchildren():
        block_tag = block_tag_by_raw.get(child.tag)
        if block_tag is None:
            continue
        match block_tag:
            case "paragraph":
                yield docx_paragraph(child, doc)
            case "table":
                yield docx_table(child, doc)
            case unreachable:
                assert_never(unreachable)


def para_text(para) -> str:
    ns_module = importlib.import_module("docx.oxml.ns")
    qn = getattr(ns_module, "qn")
    parts: list[str] = []
    for run in para.runs:
        has_image = run._r.find(qn("w:drawing")) is not None or run._r.find(qn("w:pict")) is not None
        if not has_image:
            parts.append(run.text)
    return "".join(parts).strip()


def list_depth(para) -> int:
    ns_module = importlib.import_module("docx.oxml.ns")
    qn = getattr(ns_module, "qn")
    paragraph_properties = para._p.pPr
    if paragraph_properties is None:
        return -1
    number_properties = paragraph_properties.find(qn("w:numPr"))
    if number_properties is None:
        return -1
    level = number_properties.find(qn("w:ilvl"))
    if level is None:
        return 0
    try:
        return int(level.get(qn("w:val"), 0))
    except (TypeError, ValueError):
        return 0


def _docx_cell_text(tc, text_tag: str) -> str:
    parts = [node.text for node in tc.iter() if node.tag == text_tag and node.text]
    return "".join(parts).strip()


def _docx_grid_span(tc) -> int:
    grid_span = tc.tcPr.gridSpan if tc.tcPr is not None else None
    if grid_span is None or grid_span.val is None:
        return 1
    # 손상/악의적 docx의 w:gridSpan에 매우 큰 값이 들어오면 아래 range(...) 루프가
    # 그만큼 반복돼 OOM/행업이 된다 — table_grid.py와 동일한 상한 적용.
    return min(MAX_TABLE_SPAN, max(1, int(grid_span.val)))


def _docx_vmerge_value(tc) -> str:
    vmerge = tc.tcPr.vMerge if tc.tcPr is not None else None
    return "" if vmerge is None or vmerge.val is None else str(vmerge.val)


def _docx_table_grid(table) -> tuple[list[list[str]], list[list[int]]]:
    ns_module = importlib.import_module("docx.oxml.ns")
    qn = getattr(ns_module, "qn")
    text_tag = qn("w:t")
    grid: list[list[str]] = []
    merged_cells: list[list[int]] = []
    active_vertical: dict[int, int] = {}
    for row_index, table_row in enumerate(table._tbl.tr_lst):
        row_values: list[str] = []
        col_index = 0
        for tc in table_row.tc_lst:
            while len(row_values) <= col_index:
                row_values.append("")
            col_span = _docx_grid_span(tc)
            vmerge = _docx_vmerge_value(tc)
            if vmerge == "continue":
                merge_index = active_vertical.get(col_index)
                if merge_index is not None:
                    merged_cells[merge_index][2] += 1
                row_values[col_index] = ""
            else:
                for covered_col in range(col_index, col_index + col_span):
                    active_vertical.pop(covered_col, None)
                row_values[col_index] = _docx_cell_text(tc, text_tag)
                if col_span > 1 or vmerge == "restart":
                    merged_cells.append([row_index, col_index, 1, col_span])
                    merge_index = len(merged_cells) - 1
                    if vmerge == "restart":
                        for covered_col in range(col_index, col_index + col_span):
                            active_vertical[covered_col] = merge_index
            for covered_col in range(col_index + 1, col_index + col_span):
                while len(row_values) <= covered_col:
                    row_values.append("")
                row_values[covered_col] = ""
            col_index += col_span
        if any(row_values):
            grid.append(row_values)
    max_cols = max((len(row) for row in grid), default=0)
    normalized = [row + [""] * (max_cols - len(row)) for row in grid]
    return normalized, [span for span in merged_cells if span[2] > 1 or span[3] > 1]


def parse_docx(docx_path: str) -> list[BlockDict]:
    docx_module = importlib.import_module("docx")
    table_module = importlib.import_module("docx.table")
    document_factory = getattr(docx_module, "Document")
    docx_table = getattr(table_module, "Table")

    document = document_factory(docx_path)
    blocks: list[BlockDict] = []

    for item in iter_block_items(document):
        if isinstance(item, docx_table):
            if not item.rows:
                continue
            grid, merged_cells = _docx_table_grid(item)
            header, rows = block_rows_from_grid(grid)
            if all(not header_cell for header_cell in header) and not rows:
                continue
            block: BlockDict = {"type": "table", "header": header, "rows": rows, "table_source": "docx"}
            if merged_cells:
                block["merged_cells"] = merged_cells
            blocks.append(block)
            continue

        style_name = item.style.name if item.style else ""
        text = para_text(item)

        if not text:
            continue

        heading_match = re.match(r"^(?:Heading|제목|머리말)\s*(\d+)$", style_name, re.IGNORECASE)
        if heading_match:
            level = max(1, min(int(heading_match.group(1)), 3))
            blocks.append({"type": "h", "level": level, "text": text})
            continue

        depth = list_depth(item)
        if depth >= 0:
            blocks.append({"type": "li", "text": text, "depth": min(depth, 7)})
            continue

        if re.search(r"[Qq]uote|인용", style_name):
            blocks.append({"type": "bq", "text": text})
            continue

        if re.search(r"[Cc]ode|코드", style_name):
            blocks.append({"type": "code", "text": text})
            continue

        if re.match(r"^(수신|경유|제목)\s*:", text):
            colon_idx = text.index(":")
            key = text[:colon_idx].strip()
            value = text[colon_idx + 1 :].strip()
            blocks.append({"type": "official_header", "key": key, "value": value})
            continue

        if re.search(r"[Hh]orizontal|구분선", style_name):
            blocks.append({"type": "hr"})
            continue

        blocks.append({"type": "p", "text": text})

    return blocks
