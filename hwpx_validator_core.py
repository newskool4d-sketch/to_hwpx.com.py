from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Final


HH_NS: Final = "http://www.hancom.co.kr/hwpml/2011/head"
HC_NS: Final = "http://www.hancom.co.kr/hwpml/2011/core"
HP_NS: Final = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS: Final = {"hh": HH_NS, "hc": HC_NS, "hp": HP_NS}
REQUIRED_ENTRIES: Final = ("mimetype", "Contents/header.xml")


@dataclass(frozen=True, slots=True)
class HwpxValidationIssue:
    code: str
    location: str
    message: str


@dataclass(frozen=True, slots=True)
class HwpxValidationStats:
    table_count: int
    cell_count: int
    merged_cell_count: int
    header_cell_count: int


@dataclass(frozen=True, slots=True)
class HwpxValidationReport:
    path: Path
    issues: tuple[HwpxValidationIssue, ...]
    stats: HwpxValidationStats

    @property
    def ok(self) -> bool:
        return not self.issues


def add_issue(issues: list[HwpxValidationIssue], code: str, location: str, message: str) -> None:
    issues.append(HwpxValidationIssue(code=code, location=location, message=message))


def read_xml(zf: zipfile.ZipFile, entry_name: str, issues: list[HwpxValidationIssue]) -> ET.Element | None:
    try:
        data = zf.read(entry_name)
    except KeyError:
        add_issue(issues, "missing-entry", entry_name, "required XML entry is missing")
        return None
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        add_issue(issues, "xml-parse-error", entry_name, str(exc))
        return None


def parse_positive_int(
    raw_value: str | None,
    issues: list[HwpxValidationIssue],
    location: str,
    field_name: str,
) -> int | None:
    if raw_value is None:
        add_issue(issues, "missing-attribute", location, f"{field_name} is missing")
        return None
    try:
        parsed = int(raw_value)
    except ValueError:
        add_issue(issues, "invalid-integer", location, f"{field_name} is not an integer: {raw_value}")
        return None
    if parsed <= 0:
        add_issue(issues, "invalid-positive-integer", location, f"{field_name} must be positive: {raw_value}")
        return None
    return parsed


def parse_nonnegative_int(
    raw_value: str | None,
    issues: list[HwpxValidationIssue],
    location: str,
    field_name: str,
) -> int | None:
    if raw_value is None:
        add_issue(issues, "missing-attribute", location, f"{field_name} is missing")
        return None
    try:
        parsed = int(raw_value)
    except ValueError:
        add_issue(issues, "invalid-integer", location, f"{field_name} is not an integer: {raw_value}")
        return None
    if parsed < 0:
        add_issue(issues, "invalid-nonnegative-integer", location, f"{field_name} must be nonnegative: {raw_value}")
        return None
    return parsed


def check_package_entries(zf: zipfile.ZipFile, issues: list[HwpxValidationIssue]) -> list[str]:
    names = zf.namelist()
    entries = set(names)
    for entry_name in REQUIRED_ENTRIES:
        if entry_name not in entries:
            add_issue(issues, "missing-entry", entry_name, "required package entry is missing")
    if not names or names[0] != "mimetype":
        add_issue(issues, "mimetype-order", "mimetype", "mimetype must be the first ZIP entry")
    if "mimetype" in entries:
        info = zf.getinfo("mimetype")
        if info.compress_type != zipfile.ZIP_STORED:
            add_issue(issues, "mimetype-compression", "mimetype", "mimetype must be stored without compression")
    return names


def border_fill_ids(header_root: ET.Element, issues: list[HwpxValidationIssue]) -> set[str]:
    border_fills = header_root.find(".//hh:borderFills", NS)
    if border_fills is None:
        add_issue(issues, "missing-borderfills", "Contents/header.xml", "hh:borderFills is missing")
        return set()
    border_ids: set[str] = set()
    for border_fill in border_fills.findall("hh:borderFill", NS):
        border_id = border_fill.attrib.get("id")
        if border_id is None:
            add_issue(issues, "missing-borderfill-id", "Contents/header.xml", "hh:borderFill id is missing")
            continue
        if border_id in border_ids:
            add_issue(issues, "duplicate-borderfill-id", f"Contents/header.xml#{border_id}", "duplicate borderFill id")
        border_ids.add(border_id)
    declared = parse_nonnegative_int(border_fills.attrib.get("itemCnt"), issues, "Contents/header.xml", "itemCnt")
    if declared is not None and declared != len(border_ids):
        add_issue(
            issues,
            "borderfill-count-mismatch",
            "Contents/header.xml",
            f"itemCnt={declared} but {len(border_ids)} borderFill elements exist",
        )
    return border_ids


def filled_border_fill_ids(header_root: ET.Element) -> set[str]:
    border_fills = header_root.find(".//hh:borderFills", NS)
    if border_fills is None:
        return set()
    filled_ids: set[str] = set()
    for border_fill in border_fills.findall("hh:borderFill", NS):
        border_id = border_fill.attrib.get("id")
        if border_id is None:
            continue
        brush = border_fill.find("hc:fillBrush/hc:winBrush", NS)
        if brush is not None and brush.attrib.get("faceColor"):
            filled_ids.add(border_id)
    return filled_ids
