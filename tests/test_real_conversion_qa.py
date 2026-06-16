from __future__ import annotations

import importlib.util
import zipfile
from pathlib import Path

from scripts.qa_real_conversion import QaSample
from scripts.qa_real_conversion import conversion_command
from scripts.qa_real_conversion import validate_sample_outputs
from scripts.qa_real_conversion import write_qa_report
from scripts.qa_real_conversion import write_sample_inputs


HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HS_NS = "http://www.hancom.co.kr/hwpml/2011/section"
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def _write_minimal_hwpx(path: Path) -> None:
    header = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<hh:head xmlns:hh="{HH_NS}"><hh:refList><hh:borderFills itemCnt="0"/>'
        "</hh:refList></hh:head>"
    )
    section = f'<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="{HS_NS}" xmlns:hp="{HP_NS}"/>'
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/hwp+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr("Contents/header.xml", header.encode("utf-8"))
        zf.writestr("Contents/section0.xml", section.encode("utf-8"))


def test_write_sample_inputs_creates_real_conversion_sources(tmp_path: Path) -> None:
    # Given
    if importlib.util.find_spec("openpyxl") is None:
        return

    # When
    samples = write_sample_inputs(tmp_path)

    # Then
    assert [sample.name for sample in samples] == ["minimal", "table", "merged"]
    assert all(sample.source_path.exists() for sample in samples)
    assert samples[0].source_path.suffix == ".md"
    assert samples[1].expected_tables == 1
    assert samples[2].source_path.suffix == ".xlsx"
    assert samples[2].expected_merged_cells == 1


def test_conversion_command_uses_cli_output_dir_and_timeout(tmp_path: Path) -> None:
    # Given
    sample = QaSample(
        name="table",
        source_path=tmp_path / "table.md",
        output_path=tmp_path / "out" / "table.hwpx",
        expected_tables=1,
        expected_merged_cells=0,
    )

    # When
    command = conversion_command("python", Path("to_hwpx_com.py"), sample, tmp_path / "out", 20)

    # Then
    assert command == [
        "python",
        "-B",
        "to_hwpx_com.py",
        str(sample.source_path),
        "-o",
        str(tmp_path / "out"),
        "--startup-timeout",
        "20",
    ]


def test_validate_sample_outputs_reports_missing_and_invalid_outputs(tmp_path: Path) -> None:
    # Given
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    good_output = output_dir / "minimal.hwpx"
    _write_minimal_hwpx(good_output)
    samples = (
        QaSample("minimal", tmp_path / "minimal.md", good_output, 0, 0),
        QaSample("missing", tmp_path / "missing.md", output_dir / "missing.hwpx", 1, 0),
    )

    # When
    issues = validate_sample_outputs(samples)

    # Then
    assert len(issues) == 1
    assert "missing output" in issues[0]


def test_write_qa_report_persists_summary(tmp_path: Path) -> None:
    # Given
    sample = QaSample("minimal", tmp_path / "minimal.md", tmp_path / "minimal.hwpx", 0, 0)
    from scripts.qa_real_conversion import QaRunReport

    report = QaRunReport(work_dir=tmp_path, samples=(sample,), issues=())

    # When
    report_path = write_qa_report(report)

    # Then
    assert report_path == tmp_path / "qa-report.txt"
    assert "OK: real conversion QA passed" in report_path.read_text(encoding="utf-8")
