from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import to_hwpx_com


class FakeMainHwp:
    def __init__(self) -> None:
        self.quit_called = False

    def Quit(self) -> None:
        self.quit_called = True


def test_conversion_loop_does_not_swallow_unexpected_runtime_error(tmp_path: Path, capsys) -> None:
    source = tmp_path / "bad.md"
    source.write_text("# BAD\n", encoding="utf-8")
    fake_hwp = FakeMainHwp()

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK"),
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", side_effect=RuntimeError("bad invariant")),
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        try:
            to_hwpx_com.main([str(source)])
        except RuntimeError as exc:
            assert str(exc) == "bad invariant"
        else:
            raise AssertionError("expected RuntimeError to propagate")

    captured = capsys.readouterr()
    assert fake_hwp.quit_called
    assert "bad invariant" not in captured.err
    assert "실패 목록" not in captured.err
