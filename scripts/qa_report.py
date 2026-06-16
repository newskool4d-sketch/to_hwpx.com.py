from __future__ import annotations

from pathlib import Path

from scripts.qa_models import QaRunReport


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
