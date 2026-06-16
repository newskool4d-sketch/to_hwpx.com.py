from __future__ import annotations

import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from hwpx_validator import validate_hwpx
from table_hwpx_postprocess import apply_table_width_profiles
from table_layout import TableLayout
from table_layout import table_layout_from_block


FIXTURE_PATH = Path(__file__).parent / "fixtures" / "table_corpus.json"
HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HS_NS = "http://www.hancom.co.kr/hwpml/2011/section"
NS = {"hh": HH_NS, "hc": HC_NS, "hp": HP_NS}


def _fixtures():
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def _solid_border_fill_xml(border_id: int) -> str:
    borders = "".join(
        f'<hh:{name} type="SOLID" width="0.12 mm" color="#000000"/>'
        for name in ("leftBorder", "rightBorder", "topBorder", "bottomBorder")
    )
    return (
        f'<hh:borderFill id="{border_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        f"{borders}"
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        "</hh:borderFill>"
    )


def _header_xml() -> bytes:
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


def _cell_xml(row: int, col: int) -> str:
    return (
        '<hp:tc header="0" hasMargin="0" borderFillIDRef="3">'
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
        '<hp:cellSpan colSpan="1" rowSpan="1"/>'
        '<hp:cellSz width="1" height="900"/>'
        "</hp:tc>"
    )


def _row_count(fixture) -> int:
    return len(fixture["rows"]) + 1


def _col_count(fixture) -> int:
    rows = [fixture["header"], *fixture["rows"]]
    return max(len(row) for row in rows)


def _table_xml(fixture) -> str:
    row_count = _row_count(fixture)
    col_count = _col_count(fixture)
    cells = "".join(_cell_xml(row, col) for row in range(row_count) for col in range(col_count))
    return (
        f'<hp:tbl rowCnt="{row_count}" colCnt="{col_count}" borderFillIDRef="3">'
        f'<hp:sz width="{fixture["total_width"]}"/>'
        f"{cells}</hp:tbl>"
    )


def _section_xml(fixtures) -> bytes:
    tables = "".join(_table_xml(fixture) for fixture in fixtures)
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<hs:sec xmlns:hs="{HS_NS}" xmlns:hp="{HP_NS}">'
        f"{tables}</hs:sec>"
    )
    return xml.encode("utf-8")


def _table_block(fixture):
    block = {
        "type": "table",
        "header": fixture["header"],
        "rows": fixture["rows"],
    }
    if "merged_cells" in fixture:
        block["merged_cells"] = fixture["merged_cells"]
    return block


def _write_hwpx(path: Path, fixtures) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/hwp+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr("Contents/header.xml", _header_xml())
        zf.writestr("Contents/section0.xml", _section_xml(fixtures))


def _layouts(fixtures) -> list[TableLayout]:
    return [table_layout_from_block(_table_block(fixture), total_width=fixture["total_width"]) for fixture in fixtures]


def test_table_fixture_corpus_has_expected_roles_and_widths() -> None:
    # Given
    fixtures = _fixtures()

    for fixture in fixtures:
        # When
        layout = table_layout_from_block(_table_block(fixture), total_width=fixture["total_width"])

        # Then
        assert layout.table_role == fixture["expected_role"]
        assert sum(layout.column_widths) == fixture["total_width"]
        if "expected_widths" in fixture:
            assert layout.column_widths == fixture["expected_widths"]
        if "wide_column" in fixture:
            wide_column = fixture["wide_column"]
            assert layout.column_widths[wide_column] == max(layout.column_widths)


def test_table_fixture_corpus_survives_hwpx_postprocess_and_validation(tmp_path: Path) -> None:
    # Given
    fixtures = _fixtures()
    hwpx_path = tmp_path / "table-corpus.hwpx"
    _write_hwpx(hwpx_path, fixtures)

    # When
    apply_table_width_profiles(hwpx_path, _layouts(fixtures))

    # Then
    report = validate_hwpx(hwpx_path)
    assert report.ok, [issue.code for issue in report.issues]
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        header_root = ET.fromstring(zf.read("Contents/header.xml"))
        section_root = ET.fromstring(zf.read("Contents/section0.xml"))
    header_fills = {
        border_fill.attrib["id"]
        for border_fill in header_root.findall(".//hh:borderFill", NS)
        if border_fill.find("hc:fillBrush/hc:winBrush", NS) is not None
    }
    assert header_fills
    for table_index, table in enumerate(section_root.findall(".//hp:tbl", NS)):
        fixture = fixtures[table_index]
        header_refs: set[str] = set()
        body_refs: set[str] = set()
        row_widths: dict[int, int] = {}
        for cell in table.findall(".//hp:tc", NS):
            cell_addr = cell.find("hp:cellAddr", NS)
            cell_sz = cell.find("hp:cellSz", NS)
            assert cell_addr is not None
            assert cell_sz is not None
            row = int(cell_addr.attrib["rowAddr"])
            ref = cell.attrib["borderFillIDRef"]
            if row == 0:
                header_refs.add(ref)
            else:
                body_refs.add(ref)
            row_widths[row] = row_widths.get(row, 0) + int(cell_sz.attrib["width"])
        assert header_refs <= header_fills
        assert header_refs.isdisjoint(body_refs)
        table_width = int(fixture["total_width"])
        for row, width_sum in row_widths.items():
            if "merged_cells" in fixture and row == 0:
                continue
            assert width_sum == table_width
