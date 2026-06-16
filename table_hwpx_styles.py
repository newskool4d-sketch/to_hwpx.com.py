from __future__ import annotations

import xml.etree.ElementTree as ET

from dataclasses import dataclass
from typing import Final

from table_layout import TableCellStyle


HH_NS: Final = 'http://www.hancom.co.kr/hwpml/2011/head'
HC_NS: Final = 'http://www.hancom.co.kr/hwpml/2011/core'
HP_NS: Final = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
HS_NS: Final = 'http://www.hancom.co.kr/hwpml/2011/section'
XML_NS: Final = {'hh': HH_NS, 'hc': HC_NS}
TABLE_BORDER_WIDTH: Final = '0.12 mm'
DIAGONAL_BORDER_WIDTH: Final = '0.1 mm'
TABLE_BORDER_COLOR: Final = '#000000'
BORDER_NAMES: Final = ('leftBorder', 'rightBorder', 'topBorder', 'bottomBorder')


@dataclass(frozen=True, slots=True)
class BorderFillResult:
    border_id: str
    changed: bool


@dataclass(frozen=True, slots=True)
class TableBorderFillRefs:
    body_id: str
    header_id: str
    changed: bool


def register_hwpx_namespaces() -> None:
    ET.register_namespace('hp', HP_NS)
    ET.register_namespace('hh', HH_NS)
    ET.register_namespace('hc', HC_NS)
    ET.register_namespace('hs', HS_NS)


def _rgb_color(value: int) -> str:
    return f'#{value & 0xFFFFFF:06X}'


def _border_fill_count(border_fills: ET.Element) -> int:
    return len(border_fills.findall('hh:borderFill', XML_NS))


def _next_border_fill_id(border_fills: ET.Element) -> str:
    ids: list[int] = []
    for border_fill in border_fills.findall('hh:borderFill', XML_NS):
        raw_id = border_fill.attrib.get('id')
        if raw_id is None:
            continue
        try:
            ids.append(int(raw_id))
        except ValueError:
            continue
    return str(max(ids, default=0) + 1)


def _set_border_fill_count(border_fills: ET.Element) -> None:
    border_fills.set('itemCnt', str(_border_fill_count(border_fills)))


def _border_fill_matches(border_fill: ET.Element, fill_color: str | None) -> bool:
    for border_name in BORDER_NAMES:
        border = border_fill.find(f'hh:{border_name}', XML_NS)
        if border is None:
            return False
        if border.attrib.get('type') != 'SOLID':
            return False
        if border.attrib.get('width') != TABLE_BORDER_WIDTH:
            return False
        if border.attrib.get('color') != TABLE_BORDER_COLOR:
            return False
    win_brush = border_fill.find('hc:fillBrush/hc:winBrush', XML_NS)
    if fill_color is None:
        return win_brush is None
    return win_brush is not None and win_brush.attrib.get('faceColor') == fill_color


def _make_border_fill(fill_color: str | None) -> ET.Element:
    border_fill = ET.Element(
        f'{{{HH_NS}}}borderFill',
        {
            'threeD': '0',
            'shadow': '0',
            'centerLine': 'NONE',
            'breakCellSeparateLine': '0',
        },
    )
    ET.SubElement(border_fill, f'{{{HH_NS}}}slash', {'type': 'NONE', 'Crooked': '0', 'isCounter': '0'})
    ET.SubElement(border_fill, f'{{{HH_NS}}}backSlash', {'type': 'NONE', 'Crooked': '0', 'isCounter': '0'})
    for border_name in BORDER_NAMES:
        ET.SubElement(
            border_fill,
            f'{{{HH_NS}}}{border_name}',
            {'type': 'SOLID', 'width': TABLE_BORDER_WIDTH, 'color': TABLE_BORDER_COLOR},
        )
    ET.SubElement(
        border_fill,
        f'{{{HH_NS}}}diagonal',
        {'type': 'SOLID', 'width': DIAGONAL_BORDER_WIDTH, 'color': TABLE_BORDER_COLOR},
    )
    if fill_color is not None:
        fill_brush = ET.SubElement(border_fill, f'{{{HC_NS}}}fillBrush')
        ET.SubElement(fill_brush, f'{{{HC_NS}}}winBrush', {'faceColor': fill_color, 'hatchColor': '#999999', 'alpha': '0'})
    return border_fill


def _ensure_border_fill(border_fills: ET.Element, fill_color: str | None) -> BorderFillResult:
    for border_fill in border_fills.findall('hh:borderFill', XML_NS):
        border_id = border_fill.attrib.get('id')
        if border_id is not None and _border_fill_matches(border_fill, fill_color):
            return BorderFillResult(border_id=border_id, changed=False)
    border_id = _next_border_fill_id(border_fills)
    border_fill = _make_border_fill(fill_color)
    border_fill.set('id', border_id)
    border_fills.append(border_fill)
    _set_border_fill_count(border_fills)
    return BorderFillResult(border_id=border_id, changed=True)


def ensure_table_border_fills(header_root: ET.Element, style: TableCellStyle) -> TableBorderFillRefs:
    border_fills = header_root.find('.//hh:borderFills', XML_NS)
    if border_fills is None:
        raise ValueError('header.xml borderFills not found')
    body = _ensure_border_fill(border_fills, None)
    header = _ensure_border_fill(border_fills, _rgb_color(style.header_fill_color))
    return TableBorderFillRefs(
        body_id=body.border_id,
        header_id=header.border_id,
        changed=body.changed or header.changed,
    )
