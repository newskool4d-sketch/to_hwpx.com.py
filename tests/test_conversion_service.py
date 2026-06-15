from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import conversion_service
import hwp_com
import to_hwpx_com


class FakeQuitError(RuntimeError):
    pass


class FakeConversionError(conversion_service.ConversionServiceError):
    pass


class FakeMainHwp:
    def __init__(self) -> None:
        self.quit_called = False

    def Quit(self) -> None:
        self.quit_called = True


class QuitFailingHwp(FakeMainHwp):
    def Quit(self) -> None:
        self.quit_called = True
        raise FakeQuitError("forced Quit failure")


def test_main_returns_one_and_reports_partial_failure(tmp_path: Path, capsys) -> None:
    first = tmp_path / "ok.md"
    second = tmp_path / "bad.md"
    output_dir = tmp_path / "out"
    first.write_text("# OK\n", encoding="utf-8")
    second.write_text("# BAD\n", encoding="utf-8")
    fake_hwp = FakeMainHwp()
    converted: list[str] = []

    def fake_convert_file(
        hwp,
        src_path: Path,
        hwpx_path: str,
        insert_end_mark: bool = False,
        kordoc_home: str | None = None,
    ) -> None:
        converted.append(src_path.name)
        if src_path.name == "bad.md":
            raise FakeConversionError("forced conversion failure")

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK"),
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", side_effect=fake_convert_file),
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        exit_code = to_hwpx_com.main([str(first), str(second), "-o", str(output_dir)])

    captured = capsys.readouterr()
    assert exit_code == 1
    assert converted == ["ok.md", "bad.md"]
    assert fake_hwp.quit_called
    assert "변환 중: ok.md" in captured.out
    assert "변환 중: bad.md" in captured.out
    assert "[FAIL]" in captured.err
    assert "실패 목록" in captured.err
    assert "forced conversion failure" in captured.err


def test_main_preserves_success_when_hwp_quit_fails(tmp_path: Path, capsys) -> None:
    source = tmp_path / "ok.md"
    output_dir = tmp_path / "out"
    source.write_text("# OK\n", encoding="utf-8")
    fake_hwp = QuitFailingHwp()

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK"),
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", return_value=None) as convert_file,
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        exit_code = to_hwpx_com.main([str(source), "-o", str(output_dir)])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert fake_hwp.quit_called
    convert_file.assert_called_once()
    assert "전체 변환 완료" in captured.out
    assert "HWP 종료 실패" in captured.err
    assert "forced Quit failure" in captured.err


def test_main_preserves_partial_failure_when_hwp_quit_fails(tmp_path: Path, capsys) -> None:
    source = tmp_path / "bad.md"
    output_dir = tmp_path / "out"
    source.write_text("# BAD\n", encoding="utf-8")
    fake_hwp = QuitFailingHwp()

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK"),
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", side_effect=FakeConversionError("forced conversion failure")),
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        exit_code = to_hwpx_com.main([str(source), "-o", str(output_dir)])

    captured = capsys.readouterr()
    assert exit_code == 1
    assert fake_hwp.quit_called
    assert "실패 목록" in captured.err
    assert "forced conversion failure" in captured.err
    assert "HWP 종료 실패" in captured.err
    assert "forced Quit failure" in captured.err


def test_main_returns_two_when_startup_preflight_fails(tmp_path: Path, capsys) -> None:
    source = tmp_path / "sample.md"
    source.write_text("# 제목\n", encoding="utf-8")

    with (
        patch.object(
            to_hwpx_com,
            "run_hwp_preflight",
            side_effect=hwp_com.HwpPreflightError("dispatch timeout"),
        ) as preflight,
        patch.object(to_hwpx_com, "create_hwp_object") as create_hwp_object,
        patch.object(to_hwpx_com, "convert_file") as convert_file,
    ):
        exit_code = to_hwpx_com.main([str(source), "--startup-timeout", "3"])

    captured = capsys.readouterr()
    assert exit_code == 2
    preflight.assert_called_once_with(visible=False, timeout=3)
    create_hwp_object.assert_not_called()
    convert_file.assert_not_called()
    assert "HWP 시작 사전 점검 실패" in captured.err
    assert "dispatch timeout" in captured.err


def test_main_reports_missing_input_before_hwp_startup(tmp_path: Path, capsys) -> None:
    missing = tmp_path / "missing.md"

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight") as preflight,
        patch.object(to_hwpx_com, "create_hwp_object") as create_hwp_object,
        patch.object(to_hwpx_com, "convert_file") as convert_file,
    ):
        exit_code = to_hwpx_com.main([str(missing)])

    captured = capsys.readouterr()
    assert exit_code == 1
    preflight.assert_not_called()
    create_hwp_object.assert_not_called()
    convert_file.assert_not_called()
    assert "입력 파일 없음" in captured.err
    assert str(missing) in captured.err


def test_main_reports_unsupported_extension_before_hwp_startup(tmp_path: Path, capsys) -> None:
    source = tmp_path / "sample.exe"
    source.write_text("not a document\n", encoding="utf-8")

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight") as preflight,
        patch.object(to_hwpx_com, "create_hwp_object") as create_hwp_object,
        patch.object(to_hwpx_com, "convert_file") as convert_file,
    ):
        exit_code = to_hwpx_com.main([str(source)])

    captured = capsys.readouterr()
    assert exit_code == 1
    preflight.assert_not_called()
    create_hwp_object.assert_not_called()
    convert_file.assert_not_called()
    assert "지원하지 않는 형식" in captured.err
    assert ".exe" in captured.err


def test_main_converts_valid_files_when_other_inputs_are_invalid(tmp_path: Path, capsys) -> None:
    valid = tmp_path / "ok.md"
    missing = tmp_path / "missing.md"
    output_dir = tmp_path / "out"
    valid.write_text("# OK\n", encoding="utf-8")
    fake_hwp = FakeMainHwp()
    converted: list[str] = []

    def fake_convert_file(
        hwp,
        src_path: Path,
        hwpx_path: str,
        insert_end_mark: bool = False,
        kordoc_home: str | None = None,
    ) -> None:
        converted.append(src_path.name)

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK") as preflight,
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", side_effect=fake_convert_file),
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        exit_code = to_hwpx_com.main([str(missing), str(valid), "-o", str(output_dir)])

    captured = capsys.readouterr()
    assert exit_code == 1
    preflight.assert_called_once()
    assert converted == ["ok.md"]
    assert fake_hwp.quit_called
    assert "입력 파일 없음" in captured.err
    assert "실패 목록" in captured.err


def test_preflight_cli_failure_returns_two(capsys) -> None:
    with patch.object(
        to_hwpx_com,
        "run_hwp_preflight",
        side_effect=hwp_com.HwpPreflightError("dispatch timeout"),
    ) as preflight:
        exit_code = to_hwpx_com.main(["--preflight", "--startup-timeout", "7"])

    captured = capsys.readouterr()
    assert exit_code == 2
    preflight.assert_called_once_with(visible=False, timeout=7)
    assert "[FAIL] dispatch timeout" in captured.err


def test_preflight_cli_does_not_swallow_unexpected_bug(capsys) -> None:
    # Given: the preflight path raises a programmer bug, not a typed HWP failure.
    with patch.object(to_hwpx_com, "run_hwp_preflight", side_effect=AssertionError("bad invariant")):
        # When / Then: the bug propagates instead of being reported as user input failure.
        try:
            to_hwpx_com.main(["--preflight"])
        except AssertionError as exc:
            assert str(exc) == "bad invariant"
        else:
            raise AssertionError("expected AssertionError to propagate")

    captured = capsys.readouterr()
    assert "[FAIL]" not in captured.err


def test_conversion_loop_does_not_swallow_unexpected_bug(tmp_path: Path, capsys) -> None:
    # Given: conversion enters the real loop, then convert_file raises a programmer bug.
    source = tmp_path / "bad.md"
    source.write_text("# BAD\n", encoding="utf-8")
    fake_hwp = FakeMainHwp()

    with (
        patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK"),
        patch.object(to_hwpx_com, "create_hwp_object", return_value=fake_hwp),
        patch.object(to_hwpx_com, "convert_file", side_effect=AssertionError("bad invariant")),
        patch.object(to_hwpx_com.time, "sleep", return_value=None),
    ):
        # When / Then: the bug escapes while the HWP object is still cleaned up.
        try:
            to_hwpx_com.main([str(source)])
        except AssertionError as exc:
            assert str(exc) == "bad invariant"
        else:
            raise AssertionError("expected AssertionError to propagate")

    captured = capsys.readouterr()
    assert fake_hwp.quit_called
    assert "실패 목록" not in captured.err


def test_startup_timeout_must_be_positive(capsys) -> None:
    for timeout_value in ("0", "-1", "abc"):
        error: SystemExit | None = None

        with patch.object(to_hwpx_com, "run_hwp_preflight", return_value="HWP COM preflight OK") as preflight:
            try:
                to_hwpx_com.main(["--preflight", "--startup-timeout", timeout_value])
            except SystemExit as exc:
                error = exc

        captured = capsys.readouterr()
        assert error is not None
        assert error.code == 2
        preflight.assert_not_called()
        assert "startup-timeout" in captured.err
        assert "양의 정수" in captured.err


def test_main_no_files_raises_argparse_error(capsys) -> None:
    error: SystemExit | None = None
    try:
        to_hwpx_com.main([])
    except SystemExit as exc:
        error = exc

    captured = capsys.readouterr()
    assert error is not None
    assert error.code == 2
    assert "변환할 파일 경로가 필요함" in captured.err
