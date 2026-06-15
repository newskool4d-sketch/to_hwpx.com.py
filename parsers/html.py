from __future__ import annotations

from dataclasses import dataclass
import importlib
from typing import Final, Literal, assert_never

from blocks import BlockDict

from .common import clean_inline


HtmlNodeName = Literal["h1", "h2", "h3", "p", "blockquote", "pre", "table", "li"]

SUPPORTED_HTML_NODE_NAMES: Final[tuple[HtmlNodeName, ...]] = (
    "h1",
    "h2",
    "h3",
    "p",
    "blockquote",
    "pre",
    "table",
    "li",
)
HTML_NODE_NAME_BY_RAW: Final[dict[str, HtmlNodeName]] = {
    "h1": "h1",
    "h2": "h2",
    "h3": "h3",
    "p": "p",
    "blockquote": "blockquote",
    "pre": "pre",
    "table": "table",
    "li": "li",
}


@dataclass(frozen=True, slots=True)
class MissingHtmlDependencyError(RuntimeError):
    package_name: str
    install_command: str = "pip install beautifulsoup4"

    def __str__(self) -> str:
        return f"HTML 변환에는 {self.package_name}가 필요함: {self.install_command}"


def parse_html(text: str) -> list[BlockDict]:
    try:
        bs4 = importlib.import_module("bs4")
    except ImportError as exc:
        raise MissingHtmlDependencyError(package_name="beautifulsoup4") from exc
    beautiful_soup = getattr(bs4, "BeautifulSoup")
    soup = beautiful_soup(text, "html.parser")
    blocks: list[BlockDict] = []
    body = soup.body or soup
    for node in body.find_all(SUPPORTED_HTML_NODE_NAMES, recursive=True):
        if node.find_parent(["table"]) and node.name != "table":
            continue
        if node.find_parent(["blockquote", "pre"]) and node.name not in ("blockquote", "pre"):
            continue
        if node.find_parent("li") and node.name != "li":
            continue
        node_name = HTML_NODE_NAME_BY_RAW.get(node.name)
        if node_name is None:
            continue
        match node_name:
            case "h1" | "h2" | "h3":
                value = node.get_text(" ", strip=True)
                if value:
                    blocks.append({"type": "h", "level": int(node_name[1]), "text": clean_inline(value)})
            case "p":
                value = node.get_text(" ", strip=True)
                if value:
                    blocks.append({"type": "p", "text": clean_inline(value)})
            case "blockquote":
                value = node.get_text(" ", strip=True)
                if value:
                    blocks.append({"type": "bq", "text": clean_inline(value)})
            case "pre":
                for line in node.get_text("\n", strip=True).splitlines():
                    if line.strip():
                        blocks.append({"type": "code", "text": line.rstrip()})
            case "li":
                parts: list[str] = []
                for child in node.contents:
                    if getattr(child, "name", None) in ("ul", "ol"):
                        continue
                    child_text = child.get_text(" ", strip=True) if hasattr(child, "get_text") else str(child).strip()
                    if child_text:
                        parts.append(child_text)
                value = " ".join(parts)
                if value:
                    depth = len(node.find_parents(["ul", "ol"])) - 1
                    blocks.append({"type": "li", "text": clean_inline(value), "depth": max(depth, 0)})
            case "table":
                rows: list[list[str]] = []
                for table_row in node.find_all("tr"):
                    cells = [cell.get_text(" ", strip=True) for cell in table_row.find_all(["th", "td"])]
                    if cells:
                        rows.append(cells)
                if rows:
                    blocks.append({"type": "table", "header": rows[0], "rows": rows[1:]})
            case unreachable:
                assert_never(unreachable)
    return blocks
