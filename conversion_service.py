from __future__ import annotations

import sys
import time
import uuid
from dataclasses import dataclass
from pathlib import Path

from blocks import BlockDict
from blocks import UnsupportedBlockTypeError
from converter_dispatch import detect_and_parse
from converter_dispatch import UnsupportedInputFormatError
from hwp_com import startup_exception_types
from hwp_writer import _insert_end_mark, build_doc
from parsers.html import MissingHtmlDependencyError
from parsers.pdf_errors import PdfParseError
from parsers.tabular import MissingXlsxDependencyError
from table_hwpx_postprocess import apply_table_width_profiles


class ConversionServiceError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class ConversionInputError(ConversionServiceError):
    src_path: Path
    original_error: Exception

    def __str__(self) -> str:
        return str(self.original_error)


@dataclass(frozen=True, slots=True)
class ConversionRenderError(ConversionServiceError):
    src_path: Path
    original_error: Exception

    def __str__(self) -> str:
        return str(self.original_error)


@dataclass(frozen=True, slots=True)
class HwpSaveAsError(ConversionServiceError):
    src_path: Path
    output_path: Path
    original_error: BaseException

    def __str__(self) -> str:
        return str(self.original_error)


@dataclass(frozen=True, slots=True)
class HwpxPostprocessError(ConversionServiceError):
    src_path: Path
    output_path: Path
    original_error: Exception

    def __str__(self) -> str:
        return str(self.original_error)


@dataclass(frozen=True, slots=True)
class HwpxFinalizeError(ConversionServiceError):
    src_path: Path
    output_path: Path
    original_error: OSError

    def __str__(self) -> str:
        return str(self.original_error)


INPUT_PARSE_FAILURE_TYPES: tuple[type[Exception], ...] = (
    OSError,
    UnicodeError,
    UnsupportedInputFormatError,
    MissingHtmlDependencyError,
    MissingXlsxDependencyError,
    PdfParseError,
)

POSTPROCESS_FAILURE_TYPES: tuple[type[Exception], ...] = (OSError, RuntimeError, ValueError)


def temporary_hwpx_path(final_path: Path) -> Path:
    return final_path.with_name(f".{final_path.stem}.{uuid.uuid4().hex}.tmp{final_path.suffix}")


def parse_source_blocks(src: Path, kordoc_home: str | Path | None) -> list[BlockDict]:
    try:
        return detect_and_parse(src, kordoc_home=kordoc_home)
    except INPUT_PARSE_FAILURE_TYPES as exc:
        raise ConversionInputError(src_path=src, original_error=exc) from exc


def convert_file(
    hwp,
    src_path: str | Path,
    hwpx_path: str | Path,
    insert_end_mark: bool = False,
    kordoc_home: str | Path | None = None,
) -> None:
    src = Path(src_path)
    out = Path(hwpx_path)
    if out.exists():
        raise FileExistsError(f"출력 파일이 이미 존재함: {out}")
    temp_out = temporary_hwpx_path(out)
    blocks = parse_source_blocks(src, kordoc_home)
    table_headers = [blk.get("header") or [] for blk in blocks if blk.get("type") == "table"]

    hwp.XHwpDocuments.Add(isTab=False)
    time.sleep(0.5)
    doc = hwp.XHwpDocuments.Item(hwp.XHwpDocuments.Count - 1)
    completed = False
    document_closed = False

    try:
        try:
            build_doc(hwp, blocks)
            if insert_end_mark:
                _insert_end_mark(hwp, blocks)
        except UnsupportedBlockTypeError as exc:
            raise ConversionRenderError(src_path=src, original_error=exc) from exc
        try:
            hwp.SaveAs(str(temp_out), "HWPX", "")
        except startup_exception_types() as exc:
            raise HwpSaveAsError(src_path=src, output_path=temp_out, original_error=exc) from exc
        time.sleep(0.5)
        try:
            doc.Close(isDirty=False)
            document_closed = True
        except startup_exception_types() as exc:
            document_closed = True
            print(f"[WARN] HWP 문서 닫기 실패: {exc}", file=sys.stderr)
        time.sleep(0.3)
        try:
            apply_table_width_profiles(temp_out, table_headers)
        except POSTPROCESS_FAILURE_TYPES as exc:
            raise HwpxPostprocessError(src_path=src, output_path=temp_out, original_error=exc) from exc
        try:
            temp_out.replace(out)
        except OSError as exc:
            raise HwpxFinalizeError(src_path=src, output_path=out, original_error=exc) from exc
        completed = True
    finally:
        if not document_closed:
            try:
                doc.Close(isDirty=False)
            except startup_exception_types() as exc:
                print(f"[WARN] HWP 문서 닫기 실패: {exc}", file=sys.stderr)
            time.sleep(0.3)
        if not completed and temp_out.exists():
            temp_out.unlink()
    ext = src.suffix.upper().lstrip(".")
    print(f"[완료] {ext} → {out.name}")
