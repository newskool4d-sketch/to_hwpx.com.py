from __future__ import annotations

import xml.etree.ElementTree as ET

from hwpx_validator_core import HwpxValidationIssue
from hwpx_validator_core import HwpxValidationStats
from hwpx_validator_core import NS
from hwpx_validator_core import add_issue
from hwpx_validator_core import parse_nonnegative_int
from hwpx_validator_core import parse_positive_int


def _check_border_ref(
    raw_ref: str | None,
    border_ids: set[str],
    issues: list[HwpxValidationIssue],
    location: str,
) -> None:
    if raw_ref is None:
        add_issue(issues, "missing-borderfill-ref", location, "borderFillIDRef is missing")
        return
    if raw_ref not in border_ids:
        add_issue(issues, "unknown-borderfill-ref", location, f"borderFillIDRef={raw_ref} is not defined")


def _check_cell(
    tc: ET.Element,
    border_ids: set[str],
    issues: list[HwpxValidationIssue],
    location: str,
    row_count: int | None,
    col_count: int | None,
) -> tuple[bool, bool]:
    _check_border_ref(tc.attrib.get("borderFillIDRef"), border_ids, issues, location)
    cell_addr = tc.find("hp:cellAddr", NS)
    cell_span = tc.find("hp:cellSpan", NS)
    cell_sz = tc.find("hp:cellSz", NS)
    if cell_addr is None:
        add_issue(issues, "missing-celladdr", location, "hp:cellAddr is missing")
        return False, tc.attrib.get("header") == "1"
    row = parse_nonnegative_int(cell_addr.attrib.get("rowAddr"), issues, location, "rowAddr")
    col = parse_nonnegative_int(cell_addr.attrib.get("colAddr"), issues, location, "colAddr")
    if row_count is not None and row is not None and row >= row_count:
        add_issue(issues, "celladdr-out-of-bounds", location, f"rowAddr={row} is outside rowCnt={row_count}")
    if col_count is not None and col is not None and col >= col_count:
        add_issue(issues, "celladdr-out-of-bounds", location, f"colAddr={col} is outside colCnt={col_count}")
    if cell_span is None:
        add_issue(issues, "missing-cellspan", location, "hp:cellSpan is missing")
        col_span = None
        row_span = None
    else:
        col_span = parse_positive_int(cell_span.attrib.get("colSpan"), issues, location, "colSpan")
        row_span = parse_positive_int(cell_span.attrib.get("rowSpan"), issues, location, "rowSpan")
    if cell_sz is None:
        add_issue(issues, "missing-cellsz", location, "hp:cellSz is missing")
    else:
        parse_positive_int(cell_sz.attrib.get("width"), issues, location, "width")
        parse_positive_int(cell_sz.attrib.get("height"), issues, location, "height")
    if row is not None and col is not None and row_span is not None and col_span is not None:
        if row_count is not None and row + row_span > row_count:
            add_issue(issues, "cellspan-out-of-bounds", location, f"rowSpan extends past rowCnt={row_count}")
        if col_count is not None and col + col_span > col_count:
            add_issue(issues, "cellspan-out-of-bounds", location, f"colSpan extends past colCnt={col_count}")
    return (col_span or 1) > 1 or (row_span or 1) > 1, tc.attrib.get("header") == "1"


def _attr_int(raw_value: str | None) -> int | None:
    if raw_value is None:
        return None
    try:
        return int(raw_value)
    except ValueError:
        return None


def _cell_metrics(tc: ET.Element) -> tuple[int | None, int | None, int, int]:
    cell_addr = tc.find("hp:cellAddr", NS)
    cell_span = tc.find("hp:cellSpan", NS)
    cell_sz = tc.find("hp:cellSz", NS)
    row = _attr_int(cell_addr.attrib.get("rowAddr")) if cell_addr is not None else None
    width = _attr_int(cell_sz.attrib.get("width")) if cell_sz is not None else None
    col_span = _attr_int(cell_span.attrib.get("colSpan")) if cell_span is not None else 1
    row_span = _attr_int(cell_span.attrib.get("rowSpan")) if cell_span is not None else 1
    return row, width, max(col_span or 1, 1), max(row_span or 1, 1)


def _check_row_widths(
    table_width: int | None,
    row_widths: dict[int, int],
    merged_rows: set[int],
    issues: list[HwpxValidationIssue],
    location: str,
) -> None:
    if table_width is None:
        return
    for row, width_sum in row_widths.items():
        if row in merged_rows:
            continue
        if width_sum != table_width:
            add_issue(
                issues,
                "row-width-mismatch",
                f"{location}:row[{row}]",
                f"cellSz width sum={width_sum} but table width={table_width}",
            )


def _check_table_style_refs(
    header_refs: set[str],
    body_refs: set[str],
    filled_border_ids: set[str],
    issues: list[HwpxValidationIssue],
    location: str,
) -> None:
    for border_ref in sorted(header_refs):
        if border_ref not in filled_border_ids:
            add_issue(
                issues,
                "header-borderfill-missing-fill",
                location,
                f"header borderFillIDRef={border_ref} has no fillBrush",
            )
    if header_refs and body_refs and header_refs & body_refs:
        refs = ", ".join(sorted(header_refs & body_refs))
        add_issue(issues, "header-body-borderfill-not-separated", location, f"header and body share borderFillIDRef={refs}")


def check_tables(
    section_root: ET.Element,
    section_name: str,
    border_ids: set[str],
    filled_border_ids: set[str],
    issues: list[HwpxValidationIssue],
) -> HwpxValidationStats:
    table_count = 0
    cell_count = 0
    merged_cell_count = 0
    header_cell_count = 0
    for table_index, tbl in enumerate(section_root.findall(".//hp:tbl", NS)):
        table_count += 1
        location = f"{section_name}:table[{table_index}]"
        row_count = parse_positive_int(tbl.attrib.get("rowCnt"), issues, location, "rowCnt")
        col_count = parse_positive_int(tbl.attrib.get("colCnt"), issues, location, "colCnt")
        table_sz = tbl.find("hp:sz", NS)
        table_width = (
            parse_positive_int(table_sz.attrib.get("width"), issues, location, "table width") if table_sz is not None else None
        )
        _check_border_ref(tbl.attrib.get("borderFillIDRef"), border_ids, issues, location)
        cells = tbl.findall(".//hp:tc", NS)
        if row_count is not None and col_count is not None and len(cells) != row_count * col_count:
            add_issue(issues, "cell-count-mismatch", location, f"rowCnt*colCnt={row_count * col_count} but {len(cells)} cells exist")
        seen_addrs: set[tuple[int, int]] = set()
        header_refs: set[str] = set()
        body_refs: set[str] = set()
        row_widths: dict[int, int] = {}
        merged_rows: set[int] = set()
        for cell_index, tc in enumerate(cells):
            cell_count += 1
            cell_location = f"{location}:cell[{cell_index}]"
            addr = tc.find("hp:cellAddr", NS)
            if addr is not None:
                row = parse_nonnegative_int(addr.attrib.get("rowAddr"), issues, cell_location, "rowAddr")
                col = parse_nonnegative_int(addr.attrib.get("colAddr"), issues, cell_location, "colAddr")
                if row is not None and col is not None:
                    key = (row, col)
                    if key in seen_addrs:
                        add_issue(issues, "duplicate-celladdr", cell_location, f"duplicate cell address {key}")
                    seen_addrs.add(key)
            is_merged, is_header = _check_cell(tc, border_ids, issues, cell_location, row_count, col_count)
            border_ref = tc.attrib.get("borderFillIDRef")
            if border_ref is not None:
                if is_header:
                    header_refs.add(border_ref)
                else:
                    body_refs.add(border_ref)
            row, width, col_span, row_span = _cell_metrics(tc)
            if row is not None and width is not None:
                row_widths[row] = row_widths.get(row, 0) + width
                if col_span > 1 or row_span > 1:
                    merged_rows.add(row)
            if is_merged:
                merged_cell_count += 1
            if is_header:
                header_cell_count += 1
        _check_table_style_refs(header_refs, body_refs, filled_border_ids, issues, location)
        _check_row_widths(table_width, row_widths, merged_rows, issues, location)
    return HwpxValidationStats(table_count, cell_count, merged_cell_count, header_cell_count)
