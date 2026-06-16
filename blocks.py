from __future__ import annotations

from dataclasses import dataclass
from typing import Final, TypeAlias


BlockValue: TypeAlias = str | int | list[str] | list[int] | list[list[str]] | list[list[int]]
BlockDict: TypeAlias = dict[str, BlockValue]

SUPPORTED_BLOCK_TYPES: Final = frozenset(
    {"h", "p", "li", "table", "bq", "code", "hr", "official_header"}
)


@dataclass(frozen=True, slots=True)
class UnsupportedBlockTypeError(ValueError):
    block_type: str

    def __str__(self) -> str:
        return f"지원하지 않는 block type: {self.block_type}"


def ensure_supported_block_type(block_type: str) -> str:
    if block_type not in SUPPORTED_BLOCK_TYPES:
        raise UnsupportedBlockTypeError(block_type=block_type)
    return block_type


@dataclass(frozen=True, slots=True)
class HeadingBlock:
    level: int
    text: str

    def to_dict(self) -> BlockDict:
        return {"type": "h", "level": self.level, "text": self.text}


@dataclass(frozen=True, slots=True)
class ParagraphBlock:
    text: str

    def to_dict(self) -> BlockDict:
        return {"type": "p", "text": self.text}


@dataclass(frozen=True, slots=True)
class ListItemBlock:
    text: str
    depth: int = 0
    marker: str | None = None
    content: str | None = None

    def to_dict(self) -> BlockDict:
        block: BlockDict = {"type": "li", "depth": self.depth, "text": self.text}
        if self.marker is not None:
            block["marker"] = self.marker
        if self.content is not None:
            block["content"] = self.content
        return block


@dataclass(frozen=True, slots=True)
class TableBlock:
    header: list[str]
    rows: list[list[str]]
    table_role: str | None = None
    column_widths: list[int] | None = None
    table_source: str | None = None
    worksheet_title: str | None = None
    merged_cells: list[list[int]] | None = None

    def to_dict(self) -> BlockDict:
        block: BlockDict = {
            "type": "table",
            "header": list(self.header),
            "rows": [list(row) for row in self.rows],
        }
        if self.table_role is not None:
            block["table_role"] = self.table_role
        if self.column_widths is not None:
            block["column_widths"] = list(self.column_widths)
        if self.table_source is not None:
            block["table_source"] = self.table_source
        if self.worksheet_title is not None:
            block["worksheet_title"] = self.worksheet_title
        if self.merged_cells is not None:
            block["merged_cells"] = [list(span) for span in self.merged_cells]
        return block


@dataclass(frozen=True, slots=True)
class BlockquoteBlock:
    text: str

    def to_dict(self) -> BlockDict:
        return {"type": "bq", "text": self.text}


@dataclass(frozen=True, slots=True)
class CodeBlock:
    text: str

    def to_dict(self) -> BlockDict:
        return {"type": "code", "text": self.text}


@dataclass(frozen=True, slots=True)
class HorizontalRuleBlock:
    def to_dict(self) -> BlockDict:
        return {"type": "hr"}


@dataclass(frozen=True, slots=True)
class OfficialHeaderBlock:
    key: str
    value: str

    def to_dict(self) -> BlockDict:
        return {"type": "official_header", "key": self.key, "value": self.value}


def heading(level: int, text: str) -> BlockDict:
    return HeadingBlock(level=level, text=text).to_dict()


def paragraph(text: str) -> BlockDict:
    return ParagraphBlock(text=text).to_dict()


def list_item(
    text: str,
    depth: int = 0,
    marker: str | None = None,
    content: str | None = None,
) -> BlockDict:
    return ListItemBlock(text=text, depth=depth, marker=marker, content=content).to_dict()


def table(
    header: list[str],
    rows: list[list[str]],
    table_role: str | None = None,
    column_widths: list[int] | None = None,
    table_source: str | None = None,
    worksheet_title: str | None = None,
    merged_cells: list[list[int]] | None = None,
) -> BlockDict:
    return TableBlock(
        header=header,
        rows=rows,
        table_role=table_role,
        column_widths=column_widths,
        table_source=table_source,
        worksheet_title=worksheet_title,
        merged_cells=merged_cells,
    ).to_dict()


def blockquote(text: str) -> BlockDict:
    return BlockquoteBlock(text=text).to_dict()


def code(text: str) -> BlockDict:
    return CodeBlock(text=text).to_dict()


def horizontal_rule() -> BlockDict:
    return HorizontalRuleBlock().to_dict()


def official_header(key: str, value: str) -> BlockDict:
    return OfficialHeaderBlock(key=key, value=value).to_dict()
