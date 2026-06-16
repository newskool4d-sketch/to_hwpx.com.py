from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from table_hwpx_postprocess import apply_table_width_profiles
from table_layout import table_layout_from_block


HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
NS = {"hp": HP_NS, "hh": HH_NS, "hc": HC_NS}


def _solid_border_fill_xml(border_id: int) -> str:
    borders = "".join(
        f'<hh:{name} type="SOLID" width="0.12 mm" color="#000000"/>'
        for name in ("leftBorder", "rightBorder", "topBorder", "bottomBorder")
    )
    return (
        f'<hh:borderFill id="{border_id}" threeD="0" shadow="0" '
        'centerLine="NONE" breakCellSeparateLine="0">'
        '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        f"{borders}"
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        "</hh:borderFill>"
    )


def _make_header_xml() -> bytes:
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<hh:head xmlns:hh="{HH_NS}" xmlns:hc="{HC_NS}">'
        '<hh:refList><hh:borderFills itemCnt="3">'
        '<hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"/>'
        '<hh:borderFill id="2" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"/>'
        f"{_solid_border_fill_xml(3)}"
        "</hh:borderFills></hh:refList></hh:head>"
    )
    return xml.encode("utf-8")


def _make_section_xml(col_count: int) -> bytes:
    cells: list[str] = []
    for row in range(2):
        for col in range(col_count):
            cells.append(
                f'<hp:tc header="0" hasMargin="0">'
                f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
                f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
                f'<hp:cellSz width="1" height="900"/>'
                f"</hp:tc>"
            )
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="{HP_NS}">'
        f'<hp:tbl colCnt="{col_count}" rowCnt="2"><hp:sz width="10000"/>'
        f'{"".join(cells)}</hp:tbl></hs:sec>'
    )
    return xml.encode("utf-8")


def test_postprocess_applies_layout_widths_and_cell_margins(tmp_path: Path) -> None:
    hwpx_path = tmp_path / "sample.hwpx"
    with zipfile.ZipFile(hwpx_path, "w") as zf:
        zf.writestr("Contents/header.xml", _make_header_xml())
        zf.writestr("Contents/section0.xml", _make_section_xml(5))

    layout = table_layout_from_block(
        {
            "type": "table",
            "header": ["시작", "종료", "분", "내용", "담당"],
            "rows": [["10:00", "10:05", "5", "인사", "담당자"]],
            "table_role": "schedule",
            "column_widths": [14, 14, 9, 49, 14],
        },
        total_width=10000,
    )

    apply_table_width_profiles(hwpx_path, [layout])

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        header_root = ET.fromstring(zf.read("Contents/header.xml"))
        root = ET.fromstring(zf.read("Contents/section0.xml"))
    border_fills = header_root.find(".//hh:borderFills", NS)
    assert border_fills is not None
    header_border_fill = header_root.find('.//hh:borderFill[@id="4"]', NS)
    assert header_border_fill is not None
    header_brush = header_border_fill.find("hc:fillBrush/hc:winBrush", NS)
    assert header_brush is not None
    cells = root.findall(".//hp:tc", NS)
    first_row_widths: list[int] = []
    for cell in cells:
        cell_addr = cell.find("hp:cellAddr", NS)
        cell_sz = cell.find("hp:cellSz", NS)
        assert cell_addr is not None
        assert cell_sz is not None
        if cell_addr.attrib["rowAddr"] == "0":
            first_row_widths.append(int(cell_sz.attrib["width"]))
    cell_margin = cells[-1].find("hp:cellMargin", NS)
    assert cell_margin is not None

    assert first_row_widths == [1400, 1400, 900, 4900, 1400]
    assert border_fills.attrib["itemCnt"] == "4"
    assert header_brush.attrib["faceColor"] == "#E7E7E7"
    assert cells[0].attrib["borderFillIDRef"] == "4"
    assert cells[-1].attrib["borderFillIDRef"] == "3"
    assert cells[0].attrib["header"] == "1"
    assert cells[-1].attrib["hasMargin"] == "1"
    assert cell_margin.attrib == {
        "left": "170",
        "right": "170",
        "top": "120",
        "bottom": "120",
    }


def test_postprocess_applies_merged_cell_spans(tmp_path: Path) -> None:
    hwpx_path = tmp_path / "merged.hwpx"
    with zipfile.ZipFile(hwpx_path, "w") as zf:
        zf.writestr("Contents/header.xml", _make_header_xml())
        zf.writestr("Contents/section0.xml", _make_section_xml(3))

    layout = table_layout_from_block(
        {
            "type": "table",
            "header": ["항목", "", "예산액"],
            "rows": [["강사료", "산출내역", "400,000원"]],
            "merged_cells": [[0, 0, 1, 2]],
        },
        total_width=9000,
    )

    apply_table_width_profiles(hwpx_path, [layout])

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        root = ET.fromstring(zf.read("Contents/section0.xml"))
    first_cell = root.find(".//hp:tc", NS)
    assert first_cell is not None
    cell_span = first_cell.find("hp:cellSpan", NS)
    assert cell_span is not None
    cell_sz = first_cell.find("hp:cellSz", NS)
    assert cell_sz is not None

    assert cell_span.attrib == {"colSpan": "2", "rowSpan": "1"}
    assert first_cell.attrib["borderFillIDRef"] == "4"
    assert int(cell_sz.attrib["width"]) > 3000
