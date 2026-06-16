from __future__ import annotations

import importlib
from pathlib import Path
from types import ModuleType

from hwpx_validator import format_report
from hwpx_validator import validate_hwpx
from scripts.qa_models import QaSample


class MissingQaDependencyError(RuntimeError):
    pass


def _load_openpyxl() -> ModuleType:
    try:
        return importlib.import_module("openpyxl")
    except ImportError as exc:
        raise MissingQaDependencyError("real conversion QA requires openpyxl for merged-cell XLSX input") from exc


def write_sample_inputs(work_dir: Path) -> tuple[QaSample, ...]:
    input_dir = work_dir / "inputs"
    output_dir = work_dir / "out"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    minimal = input_dir / "minimal.md"
    table = input_dir / "table.md"
    merged = input_dir / "merged.xlsx"
    minimal.write_text("# QA minimal\n\nhello\n", encoding="utf-8")
    table.write_text(
        "| 시작 | 종료 | 분 | 내용 | 담당 |\n"
        "| --- | --- | ---: | --- | --- |\n"
        "| 10:00 | 10:05 | 5 | 인사와 안내 | 홍길동 |\n"
        "| 10:05 | 10:20 | 15 | 주요 일정 공유 | 김길동 |\n",
        encoding="utf-8",
    )
    openpyxl = _load_openpyxl()
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "예산"
    worksheet.merge_cells("A1:B1")
    worksheet["A1"] = "항목"
    worksheet["C1"] = "예산액"
    worksheet.append(["강사료", "산출내역", "400,000원"])
    workbook.save(merged)
    workbook.close()
    return (
        QaSample("minimal", minimal, output_dir / "minimal.hwpx", 0, 0),
        QaSample("table", table, output_dir / "table.hwpx", 1, 0),
        QaSample("merged", merged, output_dir / "merged.hwpx", 1, 1),
    )


def conversion_command(
    python_executable: str,
    converter_path: Path,
    sample: QaSample,
    output_dir: Path,
    startup_timeout: int,
) -> list[str]:
    return [
        python_executable,
        "-B",
        str(converter_path),
        str(sample.source_path),
        "-o",
        str(output_dir),
        "--startup-timeout",
        str(startup_timeout),
    ]


def validate_sample_outputs(samples: tuple[QaSample, ...]) -> tuple[str, ...]:
    issues: list[str] = []
    for sample in samples:
        if not sample.output_path.exists():
            issues.append(f"{sample.name}: missing output {sample.output_path}")
            continue
        report = validate_hwpx(sample.output_path)
        if not report.ok:
            issues.append(f"{sample.name}: {format_report(report)}")
        if report.stats.table_count < sample.expected_tables:
            issues.append(f"{sample.name}: expected at least {sample.expected_tables} table(s), got {report.stats.table_count}")
        if report.stats.merged_cell_count < sample.expected_merged_cells:
            issues.append(
                f"{sample.name}: expected at least {sample.expected_merged_cells} merged cell(s), "
                f"got {report.stats.merged_cell_count}"
            )
    return tuple(issues)
