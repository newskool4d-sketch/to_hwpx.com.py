from __future__ import annotations

import argparse
import importlib
import os
import subprocess
import sys
import tempfile
import uuid
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Protocol


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from hwpx_validator import format_report
from hwpx_validator import validate_hwpx
from scripts.hwp_roundtrip_check import run_hwp_open_roundtrip


@dataclass(frozen=True, slots=True)
class QaSample:
    name: str
    source_path: Path
    output_path: Path
    expected_tables: int
    expected_merged_cells: int


@dataclass(frozen=True, slots=True)
class CommandResult:
    exit_code: int
    stdout: str
    stderr: str


class CommandRunner(Protocol):
    def __call__(self, command: Sequence[str]) -> CommandResult: ...


@dataclass(frozen=True, slots=True)
class QaRunReport:
    work_dir: Path
    samples: tuple[QaSample, ...]
    issues: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.issues


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


def run_subprocess_command(command: Sequence[str]) -> CommandResult:
    result = subprocess.run(
        list(command),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    return CommandResult(exit_code=result.returncode, stdout=result.stdout, stderr=result.stderr)


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


def run_real_conversion_qa(
    work_dir: Path,
    converter_path: Path,
    python_executable: str,
    startup_timeout: int,
    roundtrip_timeout: int,
    runner: CommandRunner = run_subprocess_command,
    skip_open_roundtrip: bool = False,
) -> QaRunReport:
    samples = write_sample_inputs(work_dir)
    output_dir = work_dir / "out"
    issues: list[str] = []
    for sample in samples:
        command = conversion_command(python_executable, converter_path, sample, output_dir, startup_timeout)
        print(f"[QA] converting {sample.name}: {sample.source_path.name}", flush=True)
        result = runner(command)
        if result.exit_code != 0:
            issues.append(
                f"{sample.name}: command failed exit={result.exit_code}\n"
                f"command: {' '.join(command)}\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
            )
        else:
            print(f"[QA] converted {sample.name}: {sample.output_path}", flush=True)
    issues.extend(validate_sample_outputs(samples))
    if not skip_open_roundtrip:
        merged = next(sample for sample in samples if sample.name == "merged")
        print(f"[QA] HWP open roundtrip: {merged.output_path}", flush=True)
        roundtrip_issue = run_hwp_open_roundtrip(
            python_executable,
            merged.output_path,
            work_dir / "merged-roundtrip.hwp",
            roundtrip_timeout,
        )
        if roundtrip_issue is not None:
            issues.append(roundtrip_issue)
    return QaRunReport(work_dir=work_dir, samples=samples, issues=tuple(issues))


def default_work_root() -> Path:
    if os.name == "nt":
        return Path("C:/tmp/to_hwpx_real_qa")
    return Path(tempfile.gettempdir()) / "to_hwpx_real_qa"


def new_work_dir(root: Path) -> Path:
    return root / f"run-{uuid.uuid4().hex[:12]}"


def format_qa_report(report: QaRunReport) -> str:
    lines = [f"QA work_dir={report.work_dir}", "outputs:"]
    for sample in report.samples:
        lines.append(f"- {sample.name}: {sample.output_path}")
    if report.ok:
        lines.append("OK: real conversion QA passed")
    else:
        lines.append("FAIL: real conversion QA found issues")
        lines.extend(f"- {issue}" for issue in report.issues)
    return "\n".join(lines)


def write_qa_report(report: QaRunReport, report_path: Path | None = None) -> Path:
    target = report_path or report.work_dir / "qa-report.txt"
    target.write_text(format_qa_report(report) + "\n", encoding="utf-8")
    return target


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run real HWP COM conversion QA and validate generated HWPX internals")
    parser.add_argument("--work-root", default=str(default_work_root()), help="Directory where a unique QA run folder is created")
    parser.add_argument("--converter", default=str(REPO_ROOT / "to_hwpx_com.py"))
    parser.add_argument("--python", default=sys.executable, help="Python executable used to run the converter CLI")
    parser.add_argument("--startup-timeout", type=int, default=20)
    parser.add_argument("--roundtrip-timeout", type=int, default=90)
    parser.add_argument("--skip-open-roundtrip", action="store_true", help="Skip opening merged HWPX through HWP COM")
    args = parser.parse_args(argv)
    work_dir = new_work_dir(Path(args.work_root))
    report = run_real_conversion_qa(
        work_dir=work_dir,
        converter_path=Path(args.converter),
        python_executable=args.python,
        startup_timeout=args.startup_timeout,
        roundtrip_timeout=args.roundtrip_timeout,
        skip_open_roundtrip=args.skip_open_roundtrip,
    )
    report_path = write_qa_report(report)
    print(format_qa_report(report))
    print(f"report={report_path}")
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
