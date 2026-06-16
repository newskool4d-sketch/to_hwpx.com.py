from __future__ import annotations

import sys

from scripts.qa_command import run_subprocess_command


def test_run_subprocess_command_reports_timeout() -> None:
    result = run_subprocess_command(
        [sys.executable, "-B", "-c", "import time; time.sleep(5)"],
        timeout=1,
    )

    assert result.exit_code == -1
    assert result.timed_out
