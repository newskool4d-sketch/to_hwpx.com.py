from __future__ import annotations

from pathlib import Path
import subprocess
import sys

import to_hwpx_com
from converter_dispatch import UnsupportedInputFormatError


EXPECTED_FORMATS_LINE = "지원 입력 형식: .csv, .docx, .htm, .html, .md, .pdf, .txt, .xlsx"
PROJECT_PYTHON_FILES = [
    "document_hierarchy.py",
    "hwp_writer.py",
    "hwpx_direct.py",
    "hwpx_hierarchy.py",
    "markdown_table_parser.py",
    "table_hwpx_postprocess.py",
    "table_settings.py",
    "to_hwpx_com.py",
]


def test_list_formats_outputs_sorted_supported_extensions() -> None:
    result = subprocess.run(
        [sys.executable, "-B", "to_hwpx_com.py", "--list-formats"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    assert result.returncode == 0
    assert result.stdout.strip() == EXPECTED_FORMATS_LINE
    assert result.stderr == ""


def test_project_python_files_compile() -> None:
    result = subprocess.run(
        [sys.executable, "-B", "-m", "py_compile", *PROJECT_PYTHON_FILES],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    assert result.returncode == 0, result.stderr
    assert result.stderr == ""


def test_detect_and_parse_rejects_unsupported_extension(tmp_path: Path) -> None:
    # Given
    unsupported = tmp_path / "sample.unsupported"
    unsupported.write_text("content", encoding="utf-8")

    # When
    error: UnsupportedInputFormatError | None = None
    try:
        to_hwpx_com.detect_and_parse(unsupported)
    except UnsupportedInputFormatError as exc:
        error = exc

    # Then
    assert error is not None
    assert error.extension == ".unsupported"
    assert error.supported_extensions == tuple(sorted(to_hwpx_com.SUPPORTED_EXTENSIONS))
    assert "지원하지 않는 형식" in str(error)


def test_validate_source_args_reuses_unsupported_format_error(tmp_path: Path) -> None:
    # Given
    unsupported = tmp_path / "sample.unsupported"
    unsupported.write_text("content", encoding="utf-8")

    # When
    valid_sources, failures = to_hwpx_com.validate_source_args([str(unsupported)])

    # Then
    assert valid_sources == []
    assert len(failures) == 1
    assert isinstance(failures[0][1], UnsupportedInputFormatError)
    assert "지원하지 않는 형식" in str(failures[0][1])


def test_build_output_path_uses_input_directory_and_incrementing_names(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    source.write_text("# 제목\n", encoding="utf-8")

    first = Path(to_hwpx_com.build_output_path(source, None))
    assert first == tmp_path / "sample.hwpx"

    first.write_text("existing", encoding="utf-8")
    second = Path(to_hwpx_com.build_output_path(source, None))
    assert second == tmp_path / "sample - 2.hwpx"

    second.write_text("existing", encoding="utf-8")
    third = Path(to_hwpx_com.build_output_path(source, None))
    assert third == tmp_path / "sample - 3.hwpx"


def test_build_output_path_uses_output_directory(tmp_path: Path) -> None:
    source = tmp_path / "input" / "sample.md"
    out_dir = tmp_path / "out"
    source.parent.mkdir()
    source.write_text("# 제목\n", encoding="utf-8")

    result = Path(to_hwpx_com.build_output_path(source, out_dir))

    assert result == out_dir / "sample.hwpx"
    assert out_dir.is_dir()
