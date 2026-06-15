from __future__ import annotations

from collections.abc import Callable
from collections.abc import Iterable
from enum import StrEnum
from pathlib import Path
from typing import Final
from typing import assert_never

from blocks import BlockDict
from parsers import parse_csv_file
from parsers import parse_docx
from parsers import parse_html
from parsers import parse_markdown
from parsers import parse_pdf
from parsers import parse_plain_text
from parsers import parse_xlsx


class InputFormat(StrEnum):
    MARKDOWN = ".md"
    TEXT = ".txt"
    DOCX = ".docx"
    HTML = ".html"
    HTM = ".htm"
    CSV = ".csv"
    XLSX = ".xlsx"
    PDF = ".pdf"


SUPPORTED_EXTENSIONS: Final[frozenset[str]] = frozenset(input_format.value for input_format in InputFormat)
TextParser = Callable[[str], list[BlockDict]]


class UnsupportedInputFormatError(RuntimeError):
    extension: str
    supported_extensions: tuple[str, ...]

    def __init__(self, extension: str, supported_extensions: Iterable[str]) -> None:
        self.extension = extension
        self.supported_extensions = tuple(sorted(supported_extensions))
        supported = ", ".join(self.supported_extensions)
        super().__init__(f"지원하지 않는 형식: {extension}  (지원: {supported})")


def input_format_from_extension(extension: str) -> InputFormat:
    try:
        return InputFormat(extension)
    except ValueError as exc:
        raise UnsupportedInputFormatError(extension, SUPPORTED_EXTENSIONS) from exc


def parse_text_file(path: Path, parser: TextParser) -> list[BlockDict]:
    for encoding in ("utf-8-sig", "utf-8", "cp949"):
        try:
            return parser(path.read_text(encoding=encoding))
        except UnicodeDecodeError:
            continue
    return parser(path.read_text(encoding="utf-8", errors="replace"))


def detect_and_parse(file_path: str | Path, kordoc_home: str | Path | None = None) -> list[BlockDict]:
    path = Path(file_path)
    input_format = input_format_from_extension(path.suffix.lower())
    match input_format:
        case InputFormat.MARKDOWN:
            return parse_text_file(path, parse_markdown)
        case InputFormat.TEXT:
            return parse_text_file(path, parse_plain_text)
        case InputFormat.DOCX:
            return parse_docx(str(path))
        case InputFormat.HTML | InputFormat.HTM:
            return parse_text_file(path, parse_html)
        case InputFormat.CSV:
            return parse_csv_file(path)
        case InputFormat.XLSX:
            return parse_xlsx(path)
        case InputFormat.PDF:
            return parse_pdf(path, kordoc_home=kordoc_home)
        case unreachable:
            assert_never(unreachable)
