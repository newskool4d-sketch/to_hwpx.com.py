from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import conversion_service


class ForcedCloseError(RuntimeError):
    pass


class ForcedSaveAsError(RuntimeError):
    pass


class FakeDocument:
    def __init__(self, events: list[str]) -> None:
        self.events = events

    def Close(self, isDirty: bool = False) -> None:
        self.events.append(f"Close:{isDirty}")


class CloseFailingDocument(FakeDocument):
    def Close(self, isDirty: bool = False) -> None:
        self.events.append(f"Close:{isDirty}")
        raise ForcedCloseError("forced Close failure")


class FakeDocuments:
    def __init__(self, events: list[str], fail_close: bool = False) -> None:
        self.events = events
        self.fail_close = fail_close
        self.docs: list[FakeDocument] = []

    @property
    def Count(self) -> int:
        return len(self.docs)

    def Add(self, isTab: bool = False) -> None:
        self.events.append(f"Add:{isTab}")
        document = CloseFailingDocument(self.events) if self.fail_close else FakeDocument(self.events)
        self.docs.append(document)

    def Item(self, index: int) -> FakeDocument:
        return self.docs[index]


class FakeHwp:
    def __init__(self, fail_save: bool = False, fail_close: bool = False) -> None:
        self.events: list[str] = []
        self.fail_save = fail_save
        self.XHwpDocuments = FakeDocuments(self.events, fail_close=fail_close)

    def SaveAs(self, output_path: str, format_name: str, option: str) -> None:
        self.events.append(f"SaveAs:{Path(output_path).name}:{format_name}:{option}")
        if self.fail_save:
            raise ForcedSaveAsError("forced SaveAs failure")
        Path(output_path).write_text("saved", encoding="utf-8")


def test_convert_file_closes_document_after_success(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp()

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", return_value=None),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        conversion_service.convert_file(hwp, source, output)

    assert hwp.events[0] == "Add:False"
    assert hwp.events[1].startswith("SaveAs:.sample.")
    assert hwp.events[1].endswith(".tmp.hwpx:HWPX:")
    assert hwp.events[2] == "Close:False"
    assert output.read_text(encoding="utf-8") == "saved"


def test_convert_file_closes_document_after_save_failure(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp(fail_save=True)
    error: RuntimeError | None = None

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", return_value=None),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        try:
            conversion_service.convert_file(hwp, source, output)
        except RuntimeError as exc:
            error = exc

    assert error is not None
    assert "forced SaveAs failure" in str(error)
    assert hwp.events[0] == "Add:False"
    assert hwp.events[1].startswith("SaveAs:.sample.")
    assert hwp.events[1].endswith(".tmp.hwpx:HWPX:")
    assert hwp.events[2] == "Close:False"


def test_convert_file_preserves_success_when_document_close_fails(tmp_path: Path, capsys) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp(fail_close=True)

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", return_value=None),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        conversion_service.convert_file(hwp, source, output)

    captured = capsys.readouterr()
    assert output.read_text(encoding="utf-8") == "saved"
    assert "문서 닫기 실패" in captured.err
    assert "forced Close failure" in captured.err


def test_convert_file_preserves_save_failure_when_document_close_fails(tmp_path: Path, capsys) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp(fail_save=True, fail_close=True)
    error: RuntimeError | None = None

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", return_value=None),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        try:
            conversion_service.convert_file(hwp, source, output)
        except RuntimeError as exc:
            error = exc

    captured = capsys.readouterr()
    assert error is not None
    assert "forced SaveAs failure" in str(error)
    assert "forced Close failure" not in str(error)
    assert "문서 닫기 실패" in captured.err
    assert not output.exists()
