from __future__ import annotations

import csv
import importlib
from dataclasses import dataclass
from pathlib import Path

from blocks import BlockDict

from .common import clean_inline


@dataclass(frozen=True, slots=True)
class MissingXlsxDependencyError(RuntimeError):
    package_name: str
    install_command: str = "pip install openpyxl"

    def __str__(self) -> str:
        return f"XLSX 변환에는 {self.package_name}이 필요함: {self.install_command}"


def parse_csv_file(path: str | Path) -> list[BlockDict]:
    for encoding in ("utf-8-sig", "utf-8", "cp949"):
        try:
            text = Path(path).read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = Path(path).read_text(encoding="utf-8", errors="replace")
    try:
        dialect = csv.Sniffer().sniff(text[:2048])
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(text.splitlines(), dialect))
    rows = [[clean_inline(cell.strip()) for cell in row] for row in rows if any(cell.strip() for cell in row)]
    if not rows:
        return []
    return [{"type": "table", "header": rows[0], "rows": rows[1:]}]


def parse_xlsx(path: str | Path) -> list[BlockDict]:
    try:
        openpyxl = importlib.import_module("openpyxl")
    except ImportError as exc:
        raise MissingXlsxDependencyError(package_name="openpyxl") from exc
    load_workbook = getattr(openpyxl, "load_workbook")
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        blocks: list[BlockDict] = []
        for worksheet in workbook.worksheets:
            rows: list[list[str]] = []
            for row in worksheet.iter_rows(values_only=True):
                values = ["" if value is None else clean_inline(str(value)) for value in row]
                while values and values[-1] == "":
                    values.pop()
                if any(values):
                    rows.append(values)
            if not rows:
                continue
            blocks.append({"type": "h", "level": 2, "text": worksheet.title})
            blocks.append({"type": "table", "header": rows[0], "rows": rows[1:]})
        return blocks
    finally:
        workbook.close()
