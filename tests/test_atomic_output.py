from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import conversion_service


class FakeDocument:
    def __init__(self, events: list[str]) -> None:
        self.events = events

    def Close(self, isDirty: bool = False) -> None:
        self.events.append(f"Close:{isDirty}")


class FakeDocuments:
    def __init__(self, events: list[str]) -> None:
        self.events = events
        self.docs: list[FakeDocument] = []

    @property
    def Count(self) -> int:
        return len(self.docs)

    def Add(self, isTab: bool = False) -> None:
        self.events.append(f"Add:{isTab}")
        self.docs.append(FakeDocument(self.events))

    def Item(self, index: int) -> FakeDocument:
        return self.docs[index]


class FakeHwp:
    def __init__(self) -> None:
        self.events: list[str] = []
        self.XHwpDocuments = FakeDocuments(self.events)

    def SaveAs(self, output_path: str, format_name: str, option: str) -> None:
        self.events.append(f"SaveAs:{Path(output_path).name}:{format_name}:{option}")
        Path(output_path).write_text("raw hwpx", encoding="utf-8")


def test_atomic_save_moves_postprocessed_temp_to_final(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp()
    postprocess_paths: list[Path] = []

    def fake_postprocess(path: Path, table_headers: list[list[str]]) -> None:
        postprocess_paths.append(path)
        path.write_text("postprocessed hwpx", encoding="utf-8")

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", side_effect=fake_postprocess),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        conversion_service.convert_file(hwp, source, output)

    assert output.read_text(encoding="utf-8") == "postprocessed hwpx"
    assert postprocess_paths[0].name.startswith(".sample.")
    assert postprocess_paths[0].name.endswith(".tmp.hwpx")
    assert not postprocess_paths[0].exists()
    assert list(tmp_path.glob("*.tmp.hwpx")) == []


def test_atomic_save_closes_hwp_document_before_postprocess(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp()

    def fake_postprocess(path: Path, table_headers: list[list[str]]) -> None:
        hwp.events.append("Postprocess")
        path.write_text("postprocessed hwpx", encoding="utf-8")

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", side_effect=fake_postprocess),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        conversion_service.convert_file(hwp, source, output)

    assert hwp.events[0] == "Add:False"
    assert hwp.events[1].startswith("SaveAs:.sample.")
    assert hwp.events[2] == "Close:False"
    assert hwp.events[3] == "Postprocess"
    assert output.read_text(encoding="utf-8") == "postprocessed hwpx"


def test_atomic_save_deletes_temp_when_postprocess_fails(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    hwp = FakeHwp()
    error: RuntimeError | None = None

    with (
        patch.object(conversion_service, "detect_and_parse", return_value=[{"type": "p", "text": "본문"}]),
        patch.object(conversion_service, "build_doc", return_value=None),
        patch.object(conversion_service, "apply_table_width_profiles", side_effect=RuntimeError("postprocess failed")),
        patch.object(conversion_service.time, "sleep", return_value=None),
    ):
        try:
            conversion_service.convert_file(hwp, source, output)
        except RuntimeError as exc:
            error = exc

    assert error is not None
    assert "postprocess failed" in str(error)
    assert not output.exists()
    assert list(tmp_path.glob("*.tmp.hwpx")) == []
    assert "Close:False" in hwp.events


def test_atomic_save_refuses_to_overwrite_existing_final_file(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    output = tmp_path / "sample.hwpx"
    source.write_text("# 제목\n", encoding="utf-8")
    output.write_text("previous final", encoding="utf-8")
    hwp = FakeHwp()
    error: FileExistsError | None = None

    try:
        conversion_service.convert_file(hwp, source, output)
    except FileExistsError as exc:
        error = exc

    assert error is not None
    assert output.read_text(encoding="utf-8") == "previous final"
    assert hwp.events == []
