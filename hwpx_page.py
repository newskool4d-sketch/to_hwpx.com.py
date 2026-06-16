from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Final


DEFAULT_TABLE_BODY_WIDTH: Final = 42520


def _positive_int(raw_value: str | None) -> int | None:
    if raw_value is None:
        return None
    try:
        parsed = int(raw_value)
    except ValueError:
        return None
    if parsed <= 0:
        return None
    return parsed


def _nonnegative_int(raw_value: str | None) -> int:
    parsed = _positive_int(raw_value)
    return parsed if parsed is not None else 0


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _find_descendant(root: ET.Element, name: str) -> ET.Element | None:
    for element in root.iter():
        if _local_name(element.tag) == name:
            return element
    return None


def content_width_from_secpr(secpr_xml: str, default_width: int = DEFAULT_TABLE_BODY_WIDTH) -> int:
    try:
        root = ET.fromstring(secpr_xml)
    except ET.ParseError:
        return default_width
    page_pr = root if _local_name(root.tag) == "pagePr" else _find_descendant(root, "pagePr")
    if page_pr is None:
        return default_width
    page_width = _positive_int(page_pr.attrib.get("width"))
    if page_width is None:
        return default_width
    margin = _find_descendant(page_pr, "margin")
    if margin is None:
        return page_width
    horizontal_margin = (
        _nonnegative_int(margin.attrib.get("left"))
        + _nonnegative_int(margin.attrib.get("right"))
        + _nonnegative_int(margin.attrib.get("gutter"))
    )
    content_width = page_width - horizontal_margin
    return content_width if content_width > 0 else default_width
