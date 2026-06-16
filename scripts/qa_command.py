from __future__ import annotations

import subprocess
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Protocol

from hwp_com import running_hwp_pids
from hwp_com import subprocess_timeout_text
from hwp_com import terminate_hwp_pids


@dataclass(frozen=True, slots=True)
class CommandResult:
    exit_code: int
    stdout: str
    stderr: str
    timed_out: bool = False
    cleanup_receipts: tuple[str, ...] = ()


class CommandRunner(Protocol):
    def __call__(self, command: Sequence[str], timeout: int) -> CommandResult: ...


def run_subprocess_command(command: Sequence[str], timeout: int) -> CommandResult:
    before_pids = running_hwp_pids()
    try:
        result = subprocess.run(
            list(command),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        after_pids = running_hwp_pids()
        cleanup_receipts = tuple(terminate_hwp_pids(after_pids - before_pids))
        stdout = subprocess_timeout_text(exc.stdout)
        stderr = subprocess_timeout_text(exc.stderr)
        return CommandResult(
            exit_code=-1,
            stdout=stdout,
            stderr=stderr,
            timed_out=True,
            cleanup_receipts=cleanup_receipts,
        )
    except OSError as exc:
        return CommandResult(exit_code=-1, stdout="", stderr=str(exc))
    return CommandResult(exit_code=result.returncode, stdout=result.stdout, stderr=result.stderr)
