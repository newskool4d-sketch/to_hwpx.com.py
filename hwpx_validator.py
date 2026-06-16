from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from hwpx_validator_core import HwpxValidationIssue
from hwpx_validator_core import HwpxValidationReport
from hwpx_validator_core import HwpxValidationStats
from hwpx_validator_core import add_issue
from hwpx_validator_core import border_fill_ids
from hwpx_validator_core import check_package_entries
from hwpx_validator_core import filled_border_fill_ids
from hwpx_validator_core import read_xml
from hwpx_validator_tables import check_tables


def _add_stats(left: HwpxValidationStats, right: HwpxValidationStats) -> HwpxValidationStats:
    return HwpxValidationStats(
        left.table_count + right.table_count,
        left.cell_count + right.cell_count,
        left.merged_cell_count + right.merged_cell_count,
        left.header_cell_count + right.header_cell_count,
    )


def validate_hwpx(path: str | Path) -> HwpxValidationReport:
    hwpx_path = Path(path)
    issues: list[HwpxValidationIssue] = []
    stats = HwpxValidationStats(0, 0, 0, 0)
    try:
        with zipfile.ZipFile(hwpx_path, "r") as zf:
            names = check_package_entries(zf, issues)
            header_root = read_xml(zf, "Contents/header.xml", issues)
            border_ids = border_fill_ids(header_root, issues) if header_root is not None else set()
            filled_border_ids = filled_border_fill_ids(header_root) if header_root is not None else set()
            section_names = sorted(name for name in names if name.startswith("Contents/section") and name.endswith(".xml"))
            if not section_names:
                add_issue(issues, "missing-section", "Contents", "no section XML entries were found")
            for section_name in section_names:
                section_root = read_xml(zf, section_name, issues)
                if section_root is not None:
                    stats = _add_stats(stats, check_tables(section_root, section_name, border_ids, filled_border_ids, issues))
    except zipfile.BadZipFile as exc:
        add_issue(issues, "bad-zip", str(hwpx_path), str(exc))
    except OSError as exc:
        add_issue(issues, "file-error", str(hwpx_path), str(exc))
    return HwpxValidationReport(path=hwpx_path, issues=tuple(issues), stats=stats)


def format_report(report: HwpxValidationReport) -> str:
    status = "OK" if report.ok else "FAIL"
    lines = [
        f"{status}: {report.path}",
        f"tables={report.stats.table_count} cells={report.stats.cell_count} "
        f"merged={report.stats.merged_cell_count} headers={report.stats.header_cell_count}",
    ]
    for issue in report.issues:
        lines.append(f"- [{issue.code}] {issue.location}: {issue.message}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate internal HWPX table/style structure")
    parser.add_argument("files", nargs="+", help="HWPX files to validate")
    args = parser.parse_args(argv)
    failed = False
    for file_arg in args.files:
        report = validate_hwpx(file_arg)
        print(format_report(report))
        failed = failed or not report.ok
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
