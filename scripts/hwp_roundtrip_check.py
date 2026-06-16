from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = Path("scripts/hwp_roundtrip_check.py")
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from hwp_com import HwpComError
from hwp_com import create_hwp_object
from hwp_com import running_hwp_pids
from hwp_com import startup_exception_types
from hwp_com import subprocess_timeout_text
from hwp_com import terminate_hwp_pids


def roundtrip_worker_command(python_executable: str, hwpx_path: Path, hwp_output_path: Path) -> list[str]:
    return [python_executable, "-B", str(SCRIPT_PATH), "--worker", str(hwpx_path), str(hwp_output_path)]


def _roundtrip_worker(hwpx_path: Path, hwp_output_path: Path) -> str | None:
    hwp = None
    issue: str | None = None
    try:
        hwp = create_hwp_object(visible=False)
        hwp.Open(str(hwpx_path), "HWPX", "forceopen:true")
        hwp.SaveAs(str(hwp_output_path), "HWP", "lock:false")
    except HwpComError as exc:
        issue = f"HWP open roundtrip failed: {exc}"
    except startup_exception_types() as exc:
        issue = f"HWP open roundtrip failed: {exc}"
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except startup_exception_types() as exc:
                if issue is None:
                    issue = f"HWP quit failed after roundtrip: {exc}"
    if issue is not None:
        return issue
    if not hwp_output_path.exists():
        return f"HWP open roundtrip did not create {hwp_output_path}"
    return None


def run_hwp_open_roundtrip(
    python_executable: str,
    hwpx_path: Path,
    hwp_output_path: Path,
    timeout: int,
) -> str | None:
    before_pids = running_hwp_pids()
    command = roundtrip_worker_command(python_executable, hwpx_path, hwp_output_path)
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        after_pids = running_hwp_pids()
        cleanup_receipts = terminate_hwp_pids(after_pids - before_pids)
        output = subprocess_timeout_text(exc.stdout)
        error = subprocess_timeout_text(exc.stderr)
        detail = "\n".join(part for part in (output, error) if part)
        cleanup = "\n".join(cleanup_receipts)
        suffix = f"\npartial output:\n{detail}" if detail else ""
        cleanup_suffix = f"\ncleanup:\n{cleanup}" if cleanup else ""
        return f"HWP open roundtrip timed out after {timeout} seconds.{suffix}{cleanup_suffix}"
    except OSError as exc:
        return f"HWP open roundtrip worker failed to start: {exc}"
    if result.returncode == 0:
        return None
    detail = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part.strip())
    return f"HWP open roundtrip failed exit={result.returncode}\n{detail}".strip()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Open a HWPX through HWP COM and save it as HWP")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("hwpx_path")
    parser.add_argument("hwp_output_path")
    args = parser.parse_args(argv)
    if not args.worker:
        parser.error("--worker is required")
    issue = _roundtrip_worker(Path(args.hwpx_path), Path(args.hwp_output_path))
    if issue is not None:
        print(issue, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
