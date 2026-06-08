import math
import re
import unicodedata
from typing import Dict, List, Tuple

TABLE_TOTAL_WIDTH = 14000
TABLE_MIN_ROW_HEIGHT = 900
TABLE_LINE_HEIGHT = 620
TABLE_CELL_VPAD = 260
TABLE_CELL_HPAD = 240
TABLE_UNIT_PER_VISUAL = 135

TABLE_HEADER_WIDTH_PROFILES: Dict[Tuple[str, ...], List[int]] = {
    ('구분', '내용'): [30, 70],
    ('구분', '주요 내용'): [30, 70],
    ('방향', '내용'): [30, 70],
    ('판단 사항', '검토 내용'): [30, 70],
    ('기관·부서', '역할'): [30, 70],
    ('번호', '문항', '유형'): [10, 75, 15],
    ('시간', '내용', '담당'): [20, 60, 20],
    ('단계', '내용', '시기'): [18, 62, 20],
    ('구분', '인원', '역할'): [20, 15, 65],
}

TABLE_DEFAULT_WIDTH_PROFILES = {
    2: [30, 70],
    3: [20, 40, 40],
}

COLUMN_PROFILES = {
    'index': {'min': 650, 'pref': 850, 'max': 1100, 'weight': 0.3},
    'number': {'min': 900, 'pref': 1300, 'max': 1900, 'weight': 0.6},
    'date': {'min': 1200, 'pref': 1700, 'max': 2300, 'weight': 0.7},
    'name': {'min': 850, 'pref': 1200, 'max': 1700, 'weight': 0.5},
    'position': {'min': 1000, 'pref': 1500, 'max': 2200, 'weight': 0.6},
    'org': {'min': 1600, 'pref': 2500, 'max': 3600, 'weight': 1.0},
    'title': {'min': 1700, 'pref': 2800, 'max': 4200, 'weight': 1.2},
    'detail': {'min': 2200, 'pref': 4300, 'max': 8200, 'weight': 2.3},
    'generic': {'min': 1200, 'pref': 1900, 'max': 3200, 'weight': 1.0},
}

DETAIL_HEADER_PATTERN = re.compile(
    r'(내용|세부|비고|사유|설명|의견|주소|목적|방법|추진|계획|결과|특이|주요|개요|'
    r'remark|note|description|detail|comment)',
    re.IGNORECASE,
)
ORG_HEADER_PATTERN = re.compile(r'(기관|학교|부서|소속|단체|업체|교육청|지원청|org|organization|department)', re.IGNORECASE)
TITLE_HEADER_PATTERN = re.compile(r'(명칭|제목|사업명|프로그램명|과정명|행사명|title|subject)', re.IGNORECASE)
VALUE_HEADER_PATTERN = re.compile(r'^(값|내용|value)$', re.IGNORECASE)
POSITION_HEADER_PATTERN = re.compile(r'(직위|직급|직책|보직|담당|role|position|rank)', re.IGNORECASE)
NAME_HEADER_PATTERN = re.compile(r'(성명|이름|성함|신청자|담당자|name)', re.IGNORECASE)
DATE_HEADER_PATTERN = re.compile(r'(일자|날짜|기간|시간|시각|연도|월일|date|time|period)', re.IGNORECASE)
NUMBER_HEADER_PATTERN = re.compile(r'(금액|예산|단가|합계|수량|인원|계|원|명|건|회|비율|%|amount|price|total|count|number)', re.IGNORECASE)
INDEX_HEADER_PATTERN = re.compile(r'^(순번|연번|번호|no\.?|#)$', re.IGNORECASE)
NUMBER_VALUE_PATTERN = re.compile(r'^\s*[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\s*(?:원|명|건|회|%|개|점)?\s*$')
DATE_VALUE_PATTERN = re.compile(r'^\s*\d{2,4}[./-]\d{1,2}(?:[./-]\d{1,2})?(?:\s*[-~]\s*\d{1,2}[./-]\d{1,2})?\s*$')


def _visual_width(text):
    w = 0
    for ch in str(text or ''):
        if unicodedata.combining(ch):
            continue
        if unicodedata.east_asian_width(ch) in ('F', 'W'):
            w += 2
        else:
            w += 1
    return max(w, 1)


def _normalize_table_rows(header, rows):
    all_rows = ([header] if header else []) + (rows if rows else [])
    n = max((len(row) for row in all_rows), default=0)
    normalized = []
    for row in all_rows:
        values = [str(cell or '').strip() for cell in row]
        normalized.append(values + [''] * (n - len(values)))
    return normalized, n


def _percentile(values, ratio):
    if not values:
        return 1
    ordered = sorted(values)
    idx = min(len(ordered) - 1, max(0, math.ceil(len(ordered) * ratio) - 1))
    return ordered[idx]


def _infer_col_kind(header_text, values, col_index):
    header_text = str(header_text or '').strip()
    body_values = [str(v or '').strip() for v in values if str(v or '').strip()]
    if INDEX_HEADER_PATTERN.search(header_text):
        return 'index'
    if VALUE_HEADER_PATTERN.search(header_text):
        return 'detail'
    if DETAIL_HEADER_PATTERN.search(header_text):
        return 'detail'
    if ORG_HEADER_PATTERN.search(header_text):
        return 'org'
    if TITLE_HEADER_PATTERN.search(header_text):
        return 'title'
    if POSITION_HEADER_PATTERN.search(header_text):
        return 'position'
    if NAME_HEADER_PATTERN.search(header_text):
        return 'name'
    if DATE_HEADER_PATTERN.search(header_text):
        return 'date'
    if NUMBER_HEADER_PATTERN.search(header_text):
        return 'number'
    if col_index == 0 and body_values and all(_visual_width(v) <= 4 for v in body_values):
        return 'index'
    if body_values:
        numeric_hits = sum(1 for v in body_values if NUMBER_VALUE_PATTERN.match(v))
        date_hits = sum(1 for v in body_values if DATE_VALUE_PATTERN.match(v))
        if numeric_hits / len(body_values) >= 0.75:
            return 'number'
        if date_hits / len(body_values) >= 0.6:
            return 'date'
        widths = [_visual_width(v) for v in body_values]
        avg_width = sum(widths) / len(widths)
        p90_width = _percentile(widths, 0.9)
        if p90_width >= 28 or avg_width >= 18:
            return 'detail'
        if avg_width <= 8 and all(' ' not in v for v in body_values[:10]):
            return 'name'
    return 'generic'


def _content_preferred_width(kind, header_text, values):
    widths = [_visual_width(header_text)]
    widths.extend(_visual_width(v) for v in values if str(v or '').strip())
    p90 = _percentile(widths, 0.9)
    longest = max(widths or [1])
    if kind == 'detail':
        visual_units = min(max(p90, 18), 42)
    elif kind in ('org', 'title'):
        visual_units = min(max(p90, 12), 28)
    elif kind in ('number', 'date', 'position'):
        visual_units = min(max(longest, 7), 18)
    elif kind == 'index':
        visual_units = min(max(longest, 3), 6)
    else:
        visual_units = min(max(p90, 8), 22)
    return int(visual_units * TABLE_UNIT_PER_VISUAL + TABLE_CELL_HPAD)


def _redistribute_widths(widths, kinds, total):
    if not widths:
        return []
    profiles = [COLUMN_PROFILES[k] for k in kinds]
    min_sum = sum(p['min'] for p in profiles)
    if min_sum >= total:
        result = [max(1, int(total * p['min'] / min_sum)) for p in profiles]
    else:
        result = widths[:]

    overflow = sum(result) - total
    if overflow > 0:
        shrink_room = [max(0, result[i] - profiles[i]['min']) for i in range(len(result))]
        room_sum = sum(shrink_room)
        if room_sum > 0:
            for i, room in enumerate(shrink_room):
                cut = min(room, int(overflow * room / room_sum))
                result[i] -= cut
            overflow = sum(result) - total
        while overflow > 0:
            i = max(range(len(result)), key=lambda idx: result[idx] - profiles[idx]['min'])
            if result[i] <= profiles[i]['min']:
                break
            result[i] -= 1
            overflow -= 1
    else:
        extra = total - sum(result)
        weights = [
            profiles[i]['weight'] * max(0.25, profiles[i]['max'] - result[i])
            for i in range(len(result))
        ]
        weight_sum = sum(weights)
        if weight_sum > 0:
            for i, weight in enumerate(weights):
                add = min(profiles[i]['max'] - result[i], int(extra * weight / weight_sum))
                result[i] += max(add, 0)
            extra = total - sum(result)
        while extra > 0:
            growable = [i for i, p in enumerate(profiles) if result[i] < p['max']]
            if not growable:
                break
            i = max(growable, key=lambda idx: profiles[idx]['weight'])
            result[i] += 1
            extra -= 1
        if extra > 0:
            soft_weights = [p['weight'] for p in profiles]
            soft_sum = sum(soft_weights) or len(result)
            for i, weight in enumerate(soft_weights):
                add = int(extra * weight / soft_sum)
                result[i] += add
            extra = total - sum(result)
            for i in range(extra):
                result[i % len(result)] += 1

    diff = total - sum(result)
    if diff:
        target = max(range(len(result)), key=lambda idx: profiles[idx]['weight'])
        result[target] += diff
    return result


def _profile_to_widths(profile, total=TABLE_TOTAL_WIDTH):
    if not profile:
        return []
    width_sum = sum(profile)
    if width_sum <= 0:
        return []
    widths = [max(1, int(total * value / width_sum)) for value in profile]
    diff = total - sum(widths)
    if diff:
        target = max(range(len(widths)), key=lambda idx: profile[idx])
        widths[target] += diff
    return widths


def _exact_header_profile(header):
    normalized_header = tuple(str(cell or '').strip() for cell in (header or []))
    return TABLE_HEADER_WIDTH_PROFILES.get(normalized_header)


def _default_width_profile(col_count):
    return TABLE_DEFAULT_WIDTH_PROFILES.get(col_count)


def exact_header_widths(header, col_count, total=TABLE_TOTAL_WIDTH):
    profile = _exact_header_profile(header)
    if not profile or len(profile) != col_count:
        return []
    return _profile_to_widths(profile, total)


def calc_col_widths(header, rows, total=TABLE_TOTAL_WIDTH):
    normalized, n = _normalize_table_rows(header or [], rows or [])
    if n == 0:
        return []
    if n == 1:
        return [total]
    exact_profile = _exact_header_profile(header or [])
    if exact_profile and len(exact_profile) == n:
        return _profile_to_widths(exact_profile, total)
    default_profile = _default_width_profile(n)
    if default_profile and len(default_profile) == n:
        return _profile_to_widths(default_profile, total)
    kinds = []
    preferred = []
    for ci in range(n):
        header_text = normalized[0][ci] if header else ''
        values = [row[ci] for row in normalized[1 if header else 0:]]
        kind = _infer_col_kind(header_text, values, ci)
        profile = COLUMN_PROFILES[kind]
        content_width = _content_preferred_width(kind, header_text, values)
        kinds.append(kind)
        preferred.append(max(profile['min'], min(profile['max'], max(profile['pref'], content_width))))
    return _redistribute_widths(preferred, kinds, total)


def calc_row_heights(header, rows, col_widths):
    normalized, n = _normalize_table_rows(header or [], rows or [])
    if not normalized or not col_widths:
        return []
    heights = []
    for row in normalized:
        max_lines = 1
        for ci in range(n):
            text = row[ci]
            if not text:
                continue
            usable_width = max(300, col_widths[min(ci, len(col_widths) - 1)] - TABLE_CELL_HPAD)
            capacity = max(2, int(usable_width / TABLE_UNIT_PER_VISUAL))
            visual_lines = 0
            for part in str(text).splitlines() or ['']:
                visual_lines += max(1, math.ceil(_visual_width(part) / capacity))
            max_lines = max(max_lines, visual_lines)
        heights.append(max(TABLE_MIN_ROW_HEIGHT, TABLE_CELL_VPAD + max_lines * TABLE_LINE_HEIGHT))
    return heights
