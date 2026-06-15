from __future__ import annotations

import importlib
import re
from typing import Literal, assert_never

from blocks import BlockDict


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
            header = [cell.text.strip() for cell in item.rows[0].cells]
            rows = [[cell.text.strip() for cell in row.cells] for row in item.rows[1:]]
            if all(not header_cell for header_cell in header) and not rows:
                continue
            blocks.append({"type": "table", "header": header, "rows": rows})
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
