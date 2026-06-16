from __future__ import annotations

import argparse
import os
import sys
import tempfile
import uuid
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.hwp_roundtrip_check import run_hwp_open_roundtrip
from scripts.qa_command import CommandRunner
from scripts.qa_command import run_subprocess_command
from scripts.qa_models import QaRunReport
from scripts.qa_models import QaSample
from scripts.qa_report import format_qa_report
from scripts.qa_report import write_qa_report
from scripts.qa_samples import MissingQaDependencyError
from scripts.qa_samples import conversion_command
from scripts.qa_samples import validate_sample_outputs
from scripts.qa_samples import write_sample_inputs


def run_real_conversion_qa(
    work_dir: Path,
    converter_path: Path,
    python_executable: str,
    startup_timeout: int,
    conversion_timeout: int,
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
        result = runner(command, conversion_timeout)
        if result.exit_code != 0:
            timeout_line = f"timed out after {conversion_timeout} seconds\n" if result.timed_out else ""
            cleanup = "\n".join(result.cleanup_receipts)
            cleanup_line = f"\ncleanup:\n{cleanup}" if cleanup else ""
            issues.append(
                f"{sample.name}: command failed exit={result.exit_code}\n"
                f"{timeout_line}command: {' '.join(command)}\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}{cleanup_line}"
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


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run real HWP COM conversion QA and validate generated HWPX internals")
    parser.add_argument("--work-root", default=str(default_work_root()), help="Directory where a unique QA run folder is created")
    parser.add_argument("--converter", default=str(REPO_ROOT / "to_hwpx_com.py"))
    parser.add_argument("--python", default=sys.executable, help="Python executable used to run the converter CLI")
    parser.add_argument("--startup-timeout", type=int, default=20)
    parser.add_argument("--conversion-timeout", type=int, default=120)
    parser.add_argument("--roundtrip-timeout", type=int, default=90)
    parser.add_argument("--skip-open-roundtrip", action="store_true", help="Skip opening merged HWPX through HWP COM")
    args = parser.parse_args(argv)
    work_dir = new_work_dir(Path(args.work_root))
    report = run_real_conversion_qa(
        work_dir=work_dir,
        converter_path=Path(args.converter),
        python_executable=args.python,
        startup_timeout=args.startup_timeout,
        conversion_timeout=args.conversion_timeout,
        roundtrip_timeout=args.roundtrip_timeout,
        skip_open_roundtrip=args.skip_open_roundtrip,
    )
    report_path = write_qa_report(report)
    print(format_qa_report(report))
    print(f"report={report_path}")
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
