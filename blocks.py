from __future__ import annotations

from dataclasses import dataclass
from typing import Final, TypeAlias


BlockValue: TypeAlias = str | int | list[str] | list[list[str]]
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

    def to_dict(self) -> BlockDict:
        return {
            "type": "table",
            "header": list(self.header),
            "rows": [list(row) for row in self.rows],
        }


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


def table(header: list[str], rows: list[list[str]]) -> BlockDict:
    return TableBlock(header=header, rows=rows).to_dict()


def blockquote(text: str) -> BlockDict:
    return BlockquoteBlock(text=text).to_dict()


def code(text: str) -> BlockDict:
    return CodeBlock(text=text).to_dict()


def horizontal_rule() -> BlockDict:
    return HorizontalRuleBlock().to_dict()


def official_header(key: str, value: str) -> BlockDict:
    return OfficialHeaderBlock(key=key, value=value).to_dict()
