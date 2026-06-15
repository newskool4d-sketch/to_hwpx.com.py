from __future__ import annotations

import re

from document_hierarchy import parse_hierarchy_item


def clean_inline(text: str) -> str:
    text = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", text)
    text = re.sub(r"!\[[^\]]*\]\([^\)]+\)", "", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"__([^_]+)__", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = re.sub(r"_([^_]+)_", r"\1", text)
    text = text.replace("&nbsp;", " ")
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()


def detect_list_item(line: str) -> dict[str, str | int] | None:
    item = parse_hierarchy_item(line)
    if item is None:
        return None
    content = clean_inline(item.content)
    return {
        "depth": item.depth,
        "text": f"{item.marker} {content}",
        "marker": item.marker,
        "content": content,
    }
