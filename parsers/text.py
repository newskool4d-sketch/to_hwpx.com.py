from __future__ import annotations

from blocks import BlockDict

from .common import clean_inline, detect_list_item


def parse_plain_text(text: str) -> list[BlockDict]:
    blocks: list[BlockDict] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        item = detect_list_item(line)
        if item:
            blocks.append({"type": "li", **item})
        else:
            blocks.append({"type": "p", "text": clean_inline(line)})
    return blocks
