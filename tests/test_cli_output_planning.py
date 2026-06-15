from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import to_hwpx_com


class FakeHwp:
    def __init__(self) -> None:
        self.quit_called = False

    def Quit(self) -> None:
        self.quit_called = True


def test_main_reserves_distinct_output_names_for_matching_stems(tmp_path: Path, capsys) -> None:
    left_dir = tmp_path / "left"
    right_dir = tmp_path / "right"
    output_dir = tmp_path / "out"
    left_dir.mkdir()
    right_dir.mkdir()
    left_source = left_dir / "report.md"
    right_source = right_dir / "report.md"
    left_source.write_text("# LEFT\n", encoding="utf-8")
    right_source.write_text("# RIGHT\n", encoding="utf-8")
    fake_hwp = FakeHwp()
    calls: list[tuple[str, str]] = []

    def fake_convert_file(
        hwp,
        src_path: Path,
        hwpx_path: str,
        insert_end_mark: bool = False,
        kordoc_home: str | None = None,
    ) -> None:
        calls.append((src_path.parent.name, Path(hwpx_path).name))

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK"),
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", side_effect=fake_convert_file),
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        exit_code = to_hwpx_com.main([str(left_source), str(right_source), "-o", str(output_dir)])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert calls == [("left", "report.hwpx"), ("right", "report - 2.hwpx")]
    assert fake_hwp.quit_called
    assert "전체 변환 완료" in captured.out


def test_main_reports_output_directory_file_before_hwp_startup(tmp_path: Path, capsys) -> None:
    source = tmp_path / "sample.md"
    output_path = tmp_path / "not-a-directory"
    source.write_text("# SAMPLE\n", encoding="utf-8")
    output_path.write_text("file, not a directory\n", encoding="utf-8")

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight") as preflight,
        patch.object(to_hwpx_com, "create_hwp_object") as create_hwp_object,
        patch.object(to_hwpx_com, "convert_file") as convert_file,
    ):
        exit_code = to_hwpx_com.main([str(source), "-o", str(output_path)])

    captured = capsys.readouterr()
    assert exit_code == 1
    preflight.assert_not_called()
    create_hwp_object.assert_not_called()
    convert_file.assert_not_called()
    assert "출력 경로" in captured.err


def test_plan_output_paths_reports_typed_output_preparation_error(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output_path = tmp_path / "not-a-directory"
    source.write_text("# SAMPLE\n", encoding="utf-8")
    output_path.write_text("file, not a directory\n", encoding="utf-8")

    planned_sources, failures = to_hwpx_com.plan_output_paths(
        [("sample.md", source)],
        output_path,
    )

    assert planned_sources == []
    assert len(failures) == 1
    src_arg, error = failures[0]
    assert src_arg == "sample.md"
    assert isinstance(error, to_hwpx_com.OutputPathPreparationError)
    assert error.source_arg == "sample.md"
    assert isinstance(error.original_error, OSError)
    assert "출력 경로 준비 실패" in str(error)
