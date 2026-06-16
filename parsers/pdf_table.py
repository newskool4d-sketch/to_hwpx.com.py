from __future__ import annotations

from blocks import BlockDict
from table_grid import SourceCell
from table_grid import block_rows_from_grid
from table_grid import expand_spanned_rows


def odl_cell_text(cell) -> str:
    parts: list[str] = []
    for kid in cell.get("kids", []):
        ktype = kid.get("type", "")
        if ktype in ("paragraph", "heading", "caption"):
            text = kid.get("content", "").strip()
            if text:
                parts.append(text)
        elif ktype == "text block":
            for grandkid in kid.get("kids", []):
                text = grandkid.get("content", "").strip()
                if text:
                    parts.append(text)
        elif ktype == "list":
            for item in kid.get("list items", []):
                text = item.get("content", "").strip()
                if text:
                    parts.append(text)
    return " ".join(parts)


def _odl_span(cell, keys: tuple[str, ...]) -> int:
    for key in keys:
        value = cell.get(key)
        if value is None:
            continue
        return max(1, int(value))
    return 1


def _odl_table_block(element) -> BlockDict | None:
    source_rows: list[list[SourceCell]] = []
    for row in element.get("rows", []):
        source_row = [
            SourceCell(
                text=odl_cell_text(cell),
                row_span=_odl_span(cell, ("rowspan", "rowSpan", "row span", "row_span")),
                col_span=_odl_span(cell, ("colspan", "colSpan", "col span", "col_span")),
            )
            for cell in row.get("cells", [])
        ]
        if any(cell.text for cell in source_row):
            source_rows.append(source_row)
    grid, merged_cells = expand_spanned_rows(source_rows)
    header, rows = block_rows_from_grid(grid)
    if not header:
        return None
    block: BlockDict = {"type": "table", "header": header, "rows": rows, "table_source": "pdf"}
    if merged_cells:
        block["merged_cells"] = merged_cells
    return block


def odl_element_to_blocks(element) -> list[BlockDict]:
    blocks: list[BlockDict] = []
    etype = element.get("type", "")
    if etype == "heading":
        level = min(max(int(element.get("heading level", 1)), 1), 3)
        content = element.get("content", "").strip()
        if content:
            blocks.append({"type": "h", "level": level, "text": content})
    elif etype in ("paragraph", "caption"):
        content = element.get("content", "").strip()
        if content:
            blocks.append({"type": "p", "text": content})
    elif etype == "table":
        table_block = _odl_table_block(element)
        if table_block is not None:
            blocks.append(table_block)
    elif etype == "list":
        for item in element.get("list items", []):
            content = item.get("content", "").strip()
            if content:
                blocks.append({"type": "li", "text": content, "depth": 0})
            for child in item.get("kids", []):
                blocks.extend(odl_element_to_blocks(child))
    elif etype == "text block":
        for child in element.get("kids", []):
            blocks.extend(odl_element_to_blocks(child))
    return blocks


def odl_data_to_blocks(data) -> list[BlockDict]:
    blocks: list[BlockDict] = []
    for element in data.get("kids", []):
        blocks.extend(odl_element_to_blocks(element))
    return blocks
