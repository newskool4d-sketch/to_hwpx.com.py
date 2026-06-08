import re
import unicodedata
from dataclasses import dataclass
from typing import Final


LEVEL_INDENT: Final = 1000
DEFAULT_HANGING: Final = 1200
WIDE_HANGING: Final = 1500
COM_LIST_LEFT_STEP: Final = 360
COM_TEXT_GAP: Final = 260
COM_HANG_UNIT: Final = 120

ROMAN_PATTERN: Final = re.compile(r'^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+\.|[IVXLCDM]+\.)\s+(.+)$')
ARABIC_DOT_PATTERN: Final = re.compile(r'^(\d+\.)\s+(.+)$')
HANGUL_DOT_PATTERN: Final = re.compile(r'^([가-힣]\.)\s+(.+)$')
ARABIC_PAREN_PATTERN: Final = re.compile(r'^(\d+\))\s+(.+)$')
HANGUL_PAREN_PATTERN: Final = re.compile(r'^([가-힣]\))\s+(.+)$')
PAREN_ARABIC_PATTERN: Final = re.compile(r'^(\(\d+\))\s+(.+)$')
PAREN_HANGUL_PATTERN: Final = re.compile(r'^(\([가-힣]\))\s+(.+)$')
CIRCLED_ARABIC_PATTERN: Final = re.compile(r'^([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])\s+(.+)$')
CIRCLED_HANGUL_PATTERN: Final = re.compile(r'^([㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷㉸㉹㉺㉻])\s+(.+)$')
BULLET_PATTERN: Final = re.compile(r'^([-*•])\s+(.+)$')


@dataclass(frozen=True, slots=True)
class HierarchyItem:
    depth: int
    marker: str
    content: str

    @property
    def text(self) -> str:
        return f'{self.marker} {self.content}'.strip()


@dataclass(frozen=True, slots=True)
class HierarchyStyle:
    depth: int
    left: int
    first: int
    para_pr_id: str


def visual_width(text: str) -> int:
    width = 0
    for char in text:
        if unicodedata.combining(char):
            continue
        if unicodedata.east_asian_width(char) in ('F', 'W'):
            width += 2
        else:
            width += 1
    return max(width, 1)


def parse_hierarchy_item(line: str) -> HierarchyItem | None:
    stripped = line.strip()
    patterns = (
        (0, ROMAN_PATTERN),
        (1, ARABIC_DOT_PATTERN),
        (2, HANGUL_DOT_PATTERN),
        (3, ARABIC_PAREN_PATTERN),
        (4, HANGUL_PAREN_PATTERN),
        (5, PAREN_ARABIC_PATTERN),
        (6, PAREN_HANGUL_PATTERN),
        (7, CIRCLED_ARABIC_PATTERN),
        (8, CIRCLED_HANGUL_PATTERN),
    )
    for depth, pattern in patterns:
        match = pattern.match(stripped)
        if match:
            return HierarchyItem(depth=depth, marker=match.group(1), content=match.group(2).strip())
    bullet = BULLET_PATTERN.match(stripped)
    if bullet:
        return HierarchyItem(depth=3, marker='-', content=bullet.group(2).strip())
    return None


def hwp_com_style(depth: int, marker: str) -> HierarchyStyle:
    bounded_depth = max(0, min(depth, 8))
    hanging = max(COM_TEXT_GAP, visual_width(marker or '1.') * COM_HANG_UNIT + COM_TEXT_GAP)
    left = bounded_depth * COM_LIST_LEFT_STEP + hanging
    return HierarchyStyle(depth=bounded_depth, left=left, first=-hanging, para_pr_id=hwpx_para_pr_id(bounded_depth))


def hwpx_para_metrics(depth: int) -> tuple[int, int]:
    bounded_depth = max(0, min(depth, 8))
    hanging = WIDE_HANGING if bounded_depth in (5, 6) else DEFAULT_HANGING
    marker_indent = bounded_depth * LEVEL_INDENT
    return marker_indent + hanging, -hanging


def hwpx_para_pr_id(depth: int) -> str:
    return str(200 + max(0, min(depth, 8)))
