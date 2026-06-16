from __future__ import annotations

import zipfile
from pathlib import Path

from hwpx_validator import validate_hwpx


HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HS_NS = "http://www.hancom.co.kr/hwpml/2011/section"


def _border_fill_xml(border_id: int, fill: str | None = None) -> str:
    brush = ""
    if fill is not None:
        brush = f'<hc:fillBrush><hc:winBrush faceColor="{fill}" hatchColor="#999999" alpha="0"/></hc:fillBrush>'
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
        f"{brush}"
        "</hh:borderFill>"
    )


def _header_xml(item_count: int = 4) -> bytes:
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<hh:head xmlns:hh="{HH_NS}" xmlns:hc="{HC_NS}">'
        f'<hh:refList><hh:borderFills itemCnt="{item_count}">'
        '<hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"/>'
        '<hh:borderFill id="2" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"/>'
        f"{_border_fill_xml(3)}{_border_fill_xml(4, '#E7E7E7')}"
        "</hh:borderFills></hh:refList></hh:head>"
    )
    return xml.encode("utf-8")


def _cell_xml(
    row: int,
    col: int,
    border_id: str,
    col_span: int = 1,
    row_span: int = 1,
    header: bool = False,
    width: int | None = None,
) -> str:
    header_value = "1" if header else "0"
    cell_width = width if width is not None else 2000 * col_span
    return (
        f'<hp:tc header="{header_value}" hasMargin="1" borderFillIDRef="{border_id}">'
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
        f'<hp:cellSpan colSpan="{col_span}" rowSpan="{row_span}"/>'
        f'<hp:cellSz width="{cell_width}" height="900"/>'
        "</hp:tc>"
    )


def _section_xml(extra_cell: str = "", table_width: int | None = None) -> bytes:
    cells = [
        _cell_xml(0, 0, "4", col_span=2, header=True),
        _cell_xml(0, 1, "4", header=True),
        _cell_xml(0, 2, "4", header=True),
        _cell_xml(1, 0, "3"),
        _cell_xml(1, 1, "3"),
        _cell_xml(1, 2, "3"),
    ]
    size_xml = f'<hp:sz width="{table_width}"/>' if table_width is not None else ""
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<hs:sec xmlns:hs="{HS_NS}" xmlns:hp="{HP_NS}">'
        f'<hp:tbl rowCnt="2" colCnt="3" borderFillIDRef="3">'
        f"{size_xml}"
        f'{"".join(cells)}{extra_cell}</hp:tbl>'
        "</hs:sec>"
    )
    return xml.encode("utf-8")


def _write_hwpx(path: Path, header: bytes, section: bytes) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/hwp+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr("Contents/header.xml", header)
        zf.writestr("Contents/section0.xml", section)


def test_validate_hwpx_accepts_table_border_fills_and_merged_spans(tmp_path: Path) -> None:
    # Given
    hwpx_path = tmp_path / "valid.hwpx"
    _write_hwpx(hwpx_path, _header_xml(), _section_xml())

    # When
    report = validate_hwpx(hwpx_path)

    # Then
    assert report.ok
    assert report.stats.table_count == 1
    assert report.stats.merged_cell_count == 1
    assert report.stats.header_cell_count == 3


def test_validate_hwpx_reports_header_and_border_reference_errors(tmp_path: Path) -> None:
    # Given
    hwpx_path = tmp_path / "bad-border.hwpx"
    bad_ref = _cell_xml(2, 0, "99")
    _write_hwpx(hwpx_path, _header_xml(item_count=2), _section_xml(extra_cell=bad_ref))

    # When
    report = validate_hwpx(hwpx_path)

    # Then
    assert not report.ok
    assert {"borderfill-count-mismatch", "unknown-borderfill-ref", "cell-count-mismatch"} <= {
        issue.code for issue in report.issues
    }


def test_validate_hwpx_reports_merged_cell_bounds_errors(tmp_path: Path) -> None:
    # Given
    hwpx_path = tmp_path / "bad-span.hwpx"
    bad_span = _cell_xml(1, 2, "3", col_span=2)
    _write_hwpx(hwpx_path, _header_xml(), _section_xml(extra_cell=bad_span))

    # When
    report = validate_hwpx(hwpx_path)

    # Then
    assert not report.ok
    assert "cellspan-out-of-bounds" in {issue.code for issue in report.issues}


def test_validate_hwpx_reports_header_style_contract_errors(tmp_path: Path) -> None:
    hwpx_path = tmp_path / "bad-header-style.hwpx"
    section = _section_xml().replace(b'borderFillIDRef="4"', b'borderFillIDRef="3"')
    _write_hwpx(hwpx_path, _header_xml(), section)

    report = validate_hwpx(hwpx_path)

    assert not report.ok
    assert {"header-borderfill-missing-fill", "header-body-borderfill-not-separated"} <= {
        issue.code for issue in report.issues
    }


def test_validate_hwpx_reports_unmerged_row_width_mismatch(tmp_path: Path) -> None:
    hwpx_path = tmp_path / "bad-row-width.hwpx"
    _write_hwpx(hwpx_path, _header_xml(), _section_xml(table_width=7000))

    report = validate_hwpx(hwpx_path)

    assert not report.ok
    assert "row-width-mismatch" in {issue.code for issue in report.issues}
