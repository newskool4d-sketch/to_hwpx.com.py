from __future__ import annotations

import importlib
import json
import os
import subprocess
import tempfile
from pathlib import Path

from blocks import BlockDict

from .pdf_errors import EmptyPdfTextError
from .pdf_errors import OpendataloaderPdfApiError
from .pdf_errors import OpendataloaderPdfConversionError
from .pdf_errors import OpendataloaderPdfError
from .pdf_errors import OpendataloaderPdfOutputError
from .pdf_errors import OpendataloaderPdfUnavailableError
from .pdf_errors import PdfTextExtractionError
from .pdf_errors import PdfTextFallbackError
from .text import parse_plain_text

_PDF_EXTRACTOR_ERRORS = (
    OSError,
    RuntimeError,
    ValueError,
    TypeError,
    AttributeError,
    KeyError,
    IndexError,
    UnicodeError,
)


def utf8_subprocess_env() -> dict[str, str]:
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    env.setdefault("PYTHONUTF8", "1")
    return env


def resolve_kordoc_dir(explicit_path: str | Path | None = None) -> Path | None:
    candidates: list[Path] = []
    if explicit_path:
        candidates.append(Path(explicit_path))
    for env_name in ("KORDOC_HOME", "KORDOC_AI_HOME"):
        value = os.environ.get(env_name)
        if value:
            candidates.append(Path(value))
    candidates.append(Path(r"C:/Users/홍주형/kordoc-ai"))
    for candidate in candidates:
        expanded = candidate.expanduser()
        if expanded.exists():
            return expanded.resolve()
    return None


def kordoc_commands(kordoc_dir: Path, path: str | Path) -> list[list[str]]:
    return [
        ["python", str(kordoc_dir / "main.py"), str(path)],
        ["python", "-m", "kordoc", str(path)],
    ]


def odl_cell_text(cell) -> str:
    parts: list[str] = []
    for kid in cell.get("kids", []):
        ktype = kid.get("type", "")
        if ktype in ("paragraph", "heading", "caption"):
            text = kid.get("content", "").strip()
            if text:
                parts.append(text)
        elif ktype == "text block":
            for grandkid in kid.get("kids", []):
                text = grandkid.get("content", "").strip()
                if text:
                    parts.append(text)
        elif ktype == "list":
            for item in kid.get("list items", []):
                text = item.get("content", "").strip()
                if text:
                    parts.append(text)
    return " ".join(parts)


def odl_element_to_blocks(element) -> list[BlockDict]:
    blocks: list[BlockDict] = []
    etype = element.get("type", "")
    if etype == "heading":
        level = min(max(int(element.get("heading level", 1)), 1), 3)
        content = element.get("content", "").strip()
        if content:
            blocks.append({"type": "h", "level": level, "text": content})
    elif etype in ("paragraph", "caption"):
        content = element.get("content", "").strip()
        if content:
            blocks.append({"type": "p", "text": content})
    elif etype == "table":
        grid: list[list[str]] = []
        for row in element.get("rows", []):
            row_texts = [odl_cell_text(cell) for cell in row.get("cells", [])]
            if any(row_texts):
                grid.append(row_texts)
        if grid:
            blocks.append({"type": "table", "header": grid[0], "rows": grid[1:]})
    elif etype == "list":
        for item in element.get("list items", []):
            content = item.get("content", "").strip()
            if content:
                blocks.append({"type": "li", "text": content, "depth": 0})
            for child in item.get("kids", []):
                blocks.extend(odl_element_to_blocks(child))
    elif etype == "text block":
        for child in element.get("kids", []):
            blocks.extend(odl_element_to_blocks(child))
    return blocks


def odl_data_to_blocks(data) -> list[BlockDict]:
    blocks: list[BlockDict] = []
    for element in data.get("kids", []):
        blocks.extend(odl_element_to_blocks(element))
    return blocks


def extract_pdf_blocks_odl(path: str | Path) -> list[BlockDict]:
    try:
        opendataloader_pdf = importlib.import_module("opendataloader_pdf")
    except ImportError as exc:
        raise OpendataloaderPdfUnavailableError(f"opendataloader_pdf 패키지 없음: {exc}") from exc
    try:
        convert = getattr(opendataloader_pdf, "convert")
    except AttributeError as exc:
        raise OpendataloaderPdfApiError("opendataloader_pdf.convert API 없음") from exc
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            convert(
                input_path=[str(path)],
                output_dir=tmpdir,
                format="json",
            )
        except _PDF_EXTRACTOR_ERRORS as exc:
            raise OpendataloaderPdfConversionError(f"opendataloader-pdf 변환 실패: {exc}") from exc
        json_files = sorted(Path(tmpdir).glob("*.json"))
        if not json_files:
            raise OpendataloaderPdfOutputError("opendataloader-pdf: JSON 출력 없음")
        try:
            data = json.loads(json_files[0].read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise OpendataloaderPdfOutputError(f"opendataloader-pdf JSON 읽기 실패: {exc}") from exc
    blocks = odl_data_to_blocks(data)
    if not blocks:
        raise OpendataloaderPdfOutputError("opendataloader-pdf: 추출된 blocks 없음")
    return blocks


def try_kordoc_pdf_text(
    path: str | Path,
    kordoc_home: str | Path | None = None,
    errors: list[str] | None = None,
) -> str | None:
    kordoc_dir = resolve_kordoc_dir(kordoc_home)
    if kordoc_dir is None:
        if errors is not None:
            errors.append("kordoc-ai 경로 없음")
        return None
    for cmd in kordoc_commands(kordoc_dir, path):
        try:
            result = subprocess.run(
                cmd,
                cwd=str(kordoc_dir),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=120,
                check=False,
                env=utf8_subprocess_env(),
            )
        except OSError as exc:
            if errors is not None:
                errors.append(f"kordoc-ai 실행 실패: {' '.join(cmd)}: {exc}")
            continue
        except subprocess.TimeoutExpired as exc:
            if errors is not None:
                errors.append(f"kordoc-ai 시간 초과: {' '.join(cmd)}: {exc}")
            continue
        output = (result.stdout or "").strip()
        if result.returncode == 0 and output:
            return output
        if errors is not None:
            detail = ((result.stderr or "").strip() or output or "출력 없음")
            errors.append(f"kordoc-ai 실패 exit={result.returncode}: {' '.join(cmd)}: {detail}")
    return None


def extract_pdf_text_fallback(path: str | Path) -> str:
    errors: list[str] = []
    try:
        pdfplumber = importlib.import_module("pdfplumber")
        parts: list[str] = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if text.strip():
                    parts.append(text)
        if parts:
            return "\n\n".join(parts)
    except ImportError as exc:
        errors.append(f"pdfplumber 없음: {exc}")
    except _PDF_EXTRACTOR_ERRORS as exc:
        errors.append(f"pdfplumber 실패: {exc}")
    try:
        fitz = importlib.import_module("fitz")
        fitz_open = getattr(fitz, "open")
        parts = []
        with fitz_open(str(path)) as doc:
            for page in doc:
                text = page.get_text("text") or ""
                if text.strip():
                    parts.append(text)
        if parts:
            return "\n\n".join(parts)
    except ImportError as exc:
        errors.append(f"PyMuPDF 없음: {exc}")
    except _PDF_EXTRACTOR_ERRORS as exc:
        errors.append(f"PyMuPDF 실패: {exc}")
    try:
        pypdf = importlib.import_module("pypdf")
        pdf_reader = getattr(pypdf, "PdfReader")
        reader = pdf_reader(str(path))
        parts = []
        for page in reader.pages:
            text = page.extract_text() or ""
            if text.strip():
                parts.append(text)
        if parts:
            return "\n\n".join(parts)
    except ImportError as exc:
        errors.append(f"pypdf 없음: {exc}")
    except _PDF_EXTRACTOR_ERRORS as exc:
        errors.append(f"pypdf 실패: {exc}")
    raise PdfTextFallbackError("PDF 텍스트 추출 실패: " + ("; ".join(errors) or "사용 가능한 추출기 없음"))


def parse_pdf(path: str | Path, kordoc_home: str | Path | None = None) -> list[BlockDict]:
    try:
        return extract_pdf_blocks_odl(path)
    except OpendataloaderPdfError as odl_exc:
        odl_warn = str(odl_exc)

    kordoc_errors: list[str] = []
    text = try_kordoc_pdf_text(path, kordoc_home=kordoc_home, errors=kordoc_errors)
    if text is None:
        try:
            text = extract_pdf_text_fallback(path)
        except PdfTextFallbackError as fb_exc:
            kordoc_detail = "; ".join(kordoc_errors) if kordoc_errors else "시도 안 함"
            raise PdfTextExtractionError(
                f"PDF 텍스트 추출 실패.\n  opendataloader-pdf: {odl_warn}\n  kordoc-ai: {kordoc_detail}\n  fallback: {fb_exc}"
            ) from fb_exc

    if not text.strip():
        raise EmptyPdfTextError(f"PDF에서 텍스트를 추출하지 못함: {path}")
    return parse_plain_text(text)
