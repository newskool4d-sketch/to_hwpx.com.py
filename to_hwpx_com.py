"""
HWP COM 자동화로 Markdown / DOCX → HWPX 변환 (v1)
확장자(.md / .docx)를 자동 감지하여 적절한 파서로 처리.
이미지는 skip.

변경 이력:
  v1 - md_to_hwpx_com v3 + docx_to_hwpx_com v1 통합
"""
import win32com.client
import math
import re
import os
import time
import unicodedata
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET


# ─── Markdown 파서 ─────────────────────────────────────────────────────────────

def _clean_inline(text):
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'!\[[^\]]*\]\([^\)]+\)', '', text)
    text = re.sub(r'`([^`]+)`', r'\1', text)
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
    text = re.sub(r'__([^_]+)__', r'\1', text)
    text = re.sub(r'\*([^*]+)\*', r'\1', text)
    text = re.sub(r'_([^_]+)_', r'\1', text)
    text = text.replace('&nbsp;', ' ')
    text = re.sub(r'<[^>]+>', '', text)
    return text.strip()

def _is_separator(line):
    if len(line) > 500:
        return False
    cells = line.strip().strip('|').split('|')
    return len(cells) >= 1 and all(re.match(r'^[ \t]*:?-+:?[ \t]*$', c) for c in cells)

def _parse_table_row(line):
    line = line.strip().strip('|')
    return [_clean_inline(c.strip()) for c in line.split('|')]

def _detect_list_item(line):
    """
    한국 행정문서 8단계 항목 체계 감지.
    Returns (depth, display_text) or None.
    depth: 0=1./- , 1=가., 2=1), 3=가), 4=(1), 5=(가), 6=①, 7=㉮
    """
    stripped = line.strip()
    checks = [
        (7, re.compile(r'^([㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷])\s+(.*)')),
        (6, re.compile(r'^([①②③④⑤⑥⑦⑧⑨⑩])\s+(.*)')),
        (5, re.compile(r'^(\([가나다라마바사아자차카타파하]\))\s+(.*)')),
        (4, re.compile(r'^(\(\d+\))\s+(.*)')),
        (3, re.compile(r'^([가나다라마바사아자차카타파하]\))\s+(.*)')),
        (2, re.compile(r'^(\d+\))\s+(.*)')),
        (1, re.compile(r'^([가나다라마바사아자차카타파하]\.)\s+(.*)')),
        (0, re.compile(r'^(\d+\.)\s+(.*)')),
    ]
    for depth, pattern in checks:
        m = pattern.match(stripped)
        if m:
            marker = m.group(1)
            content = _clean_inline(m.group(2))
            return (depth, f'{marker} {content}')
    m = re.match(r'^[-*]\s+(.*)', stripped)
    if m:
        return (0, '• ' + _clean_inline(m.group(1)))
    return None


def parse_markdown(text):
    lines = text.splitlines()
    blocks = []
    i = 0
    in_front = False

    while i < len(lines):
        line = lines[i]

        if not line.strip():
            i += 1
            continue

        # frontmatter skip
        if line.strip() == '---':
            if i == 0:
                in_front = True
                i += 1
                continue
            elif in_front:
                in_front = False
                i += 1
                continue
            else:
                blocks.append({'type': 'hr'})
                i += 1
                continue

        if in_front:
            i += 1
            continue

        stripped_line = line.strip()

        # 공문 헤더 (수신/경유/제목)
        if re.match(r'^(수신|경유|제목)\s*:', stripped_line):
            colon_idx = stripped_line.index(':')
            key = stripped_line[:colon_idx].strip()
            value = _clean_inline(stripped_line[colon_idx + 1:].strip())
            blocks.append({'type': 'official_header', 'key': key, 'value': value})
            i += 1
            continue

        # HR
        if re.match(r'^-{3,}\s*$', line) or re.match(r'^\*{3,}\s*$', line):
            blocks.append({'type': 'hr'})
            i += 1
            continue

        # 제목
        m = re.match(r'^(#{1,3})\s+(.*)', line)
        if m:
            blocks.append({'type': 'h', 'level': len(m.group(1)), 'text': _clean_inline(m.group(2))})
            i += 1
            continue

        # 표
        if line.strip().startswith('|') and i + 1 < len(lines) and _is_separator(lines[i + 1]):
            header = _parse_table_row(line)
            i += 2
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append(_parse_table_row(lines[i]))
                i += 1
            blocks.append({'type': 'table', 'header': header, 'rows': rows})
            continue

        # 항목 체계 (8단계)
        li_result = _detect_list_item(line)
        if li_result:
            depth, text = li_result
            blocks.append({'type': 'li', 'text': text, 'depth': depth})
            i += 1
            continue

        # blockquote
        if line.strip().startswith('>'):
            text = re.sub(r'^>\s*', '', line.strip())
            if text:
                blocks.append({'type': 'bq', 'text': _clean_inline(text)})
            i += 1
            continue

        # 코드블록
        if line.strip().startswith('```'):
            i += 1
            code_lines = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1
            for cl in code_lines:
                if cl.strip():
                    blocks.append({'type': 'code', 'text': cl})
            continue

        # 일반 단락
        t = _clean_inline(line)
        if t:
            blocks.append({'type': 'p', 'text': t})
        i += 1

    return blocks


# ─── DOCX 파서 ─────────────────────────────────────────────────────────────────

def _iter_block_items(doc):
    """문서 본문의 단락·표를 원래 순서대로 yield."""
    from docx.oxml.ns import qn
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph

    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield DocxParagraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield DocxTable(child, doc)


def _para_text(para):
    """단락의 전체 텍스트. 이미지 run은 skip."""
    from docx.oxml.ns import qn
    parts = []
    for run in para.runs:
        has_image = (
            run._r.find(qn('w:drawing')) is not None
            or run._r.find(qn('w:pict')) is not None
        )
        if not has_image:
            parts.append(run.text)
    return ''.join(parts).strip()


def _list_depth(para):
    """목록 들여쓰기 레벨(0-based). 목록 아니면 -1."""
    from docx.oxml.ns import qn
    pPr = para._p.pPr
    if pPr is None:
        return -1
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        return -1
    ilvl = numPr.find(qn('w:ilvl'))
    if ilvl is None:
        return 0
    try:
        return int(ilvl.get(qn('w:val'), 0))
    except (TypeError, ValueError):
        return 0


def parse_docx(docx_path):
    from docx import Document
    from docx.table import Table as DocxTable

    doc = Document(docx_path)
    blocks = []

    for item in _iter_block_items(doc):

        # 표
        if isinstance(item, DocxTable):
            if not item.rows:
                continue
            header = [cell.text.strip() for cell in item.rows[0].cells]
            rows = [
                [cell.text.strip() for cell in row.cells]
                for row in item.rows[1:]
            ]
            if all(not h for h in header) and not rows:
                continue
            blocks.append({'type': 'table', 'header': header, 'rows': rows})
            continue

        # 단락
        para = item
        style_name = para.style.name if para.style else ''
        text = _para_text(para)

        if not text:
            continue

        # 제목
        heading_match = re.match(
            r'^(?:Heading|제목|머리말)\s*(\d+)$', style_name, re.IGNORECASE
        )
        if heading_match:
            level = max(1, min(int(heading_match.group(1)), 3))
            blocks.append({'type': 'h', 'level': level, 'text': text})
            continue

        # 목록
        depth = _list_depth(para)
        if depth >= 0:
            blocks.append({'type': 'li', 'text': text, 'depth': min(depth, 7)})
            continue

        # 인용
        if re.search(r'[Qq]uote|인용', style_name):
            blocks.append({'type': 'bq', 'text': text})
            continue

        # 코드
        if re.search(r'[Cc]ode|코드', style_name):
            blocks.append({'type': 'code', 'text': text})
            continue

        # 공문 헤더
        if re.match(r'^(수신|경유|제목)\s*:', text):
            colon_idx = text.index(':')
            key = text[:colon_idx].strip()
            value = text[colon_idx + 1:].strip()
            blocks.append({'type': 'official_header', 'key': key, 'value': value})
            continue

        # 수평선
        if re.search(r'[Hh]orizontal|구분선', style_name):
            blocks.append({'type': 'hr'})
            continue

        # 일반 단락
        blocks.append({'type': 'p', 'text': text})

    return blocks


# ─── 확장자 자동 감지 ──────────────────────────────────────────────────────────

def detect_and_parse(file_path):
    """확장자에 따라 적절한 파서를 선택하여 블록 리스트 반환."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.md':
        with open(file_path, 'r', encoding='utf-8') as f:
            return parse_markdown(f.read())
    elif ext == '.docx':
        return parse_docx(file_path)
    else:
        raise ValueError(f'지원하지 않는 형식: {ext}  (.md 또는 .docx만 가능)')


# ─── HWP COM 헬퍼 ─────────────────────────────────────────────────────────────

def insert_text(hwp, text):
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)

def break_para(hwp):
    hwp.HAction.Run('BreakPara')

def set_char_shape(hwp, height=1300, bold=False, italic=False, font='body'):
    face_hangul = '휴먼명조' if font == 'body' else '맑은 고딕'
    face_latin  = 'Arial'
    act = hwp.CreateAction('CharShape')
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem('Height', height)
    pset.SetItem('Bold', bold)
    pset.SetItem('Italic', italic)
    pset.SetItem('FaceNameHangul', face_hangul)
    pset.SetItem('FaceNameLatin', face_latin)
    act.Execute(pset)

def set_para_shape(hwp, align=0, space_before=0, space_after=0,
                   indent_left=0, indent_first=0):
    act = hwp.CreateAction('ParagraphShape')
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem('Align', align)
    pset.SetItem('SpaceBefore', space_before)
    pset.SetItem('SpaceAfter', space_after)
    pset.SetItem('IndentLeft', indent_left)
    pset.SetItem('IndentFirst', indent_first)
    act.Execute(pset)


# ─── 표 열 너비 산정 ───────────────────────────────────────────────────────────

TABLE_TOTAL_WIDTH = 14000
TABLE_MIN_ROW_HEIGHT = 900
TABLE_LINE_HEIGHT = 620
TABLE_CELL_VPAD = 260
TABLE_CELL_HPAD = 240
TABLE_UNIT_PER_VISUAL = 135

TABLE_HEADER_WIDTH_PROFILES = {
    ('구분', '내용'): [25, 75],
    ('구분', '주요 내용'): [25, 75],
    ('방향', '내용'): [25, 75],
    ('판단 사항', '검토 내용'): [30, 70],
    ('기관·부서', '역할'): [30, 70],
    ('번호', '문항', '유형'): [10, 75, 15],
    ('시간', '내용', '담당'): [20, 60, 20],
    ('단계', '내용', '시기'): [18, 62, 20],
    ('구분', '인원', '역할'): [20, 15, 65],
}

TABLE_DEFAULT_WIDTH_PROFILES = {
    2: [28, 72],
    3: [18, 62, 20],
    4: [15, 35, 25, 25],
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


def _table_header_profile(header, col_count):
    normalized_header = tuple(str(cell or '').strip() for cell in (header or []))
    if normalized_header in TABLE_HEADER_WIDTH_PROFILES:
        return TABLE_HEADER_WIDTH_PROFILES[normalized_header]
    if col_count in TABLE_DEFAULT_WIDTH_PROFILES:
        return TABLE_DEFAULT_WIDTH_PROFILES[col_count]
    return None


def calc_col_widths(header, rows, total=TABLE_TOTAL_WIDTH):
    normalized, n = _normalize_table_rows(header or [], rows or [])
    if n == 0:
        return []
    if n == 1:
        return [total]
    profile = _table_header_profile(header or [], n)
    if profile and len(profile) == n:
        return _profile_to_widths(profile, total)
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


def _extract_plain_text(elem):
    return ''.join(elem.itertext()).strip()


def _rewrite_zip_entry(zip_path, entry_name, data):
    src = os.fspath(zip_path)
    fd, tmp_name = tempfile.mkstemp(suffix='.hwpx')
    os.close(fd)
    try:
        with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(tmp_name, 'w') as zout:
            for item in zin.infolist():
                content = data if item.filename == entry_name else zin.read(item.filename)
                zi = zipfile.ZipInfo(item.filename, item.date_time)
                zi.comment = item.comment
                zi.extra = item.extra
                zi.internal_attr = item.internal_attr
                zi.external_attr = item.external_attr
                zi.create_system = item.create_system
                zi.compress_type = item.compress_type
                zout.writestr(zi, content)
        shutil.move(tmp_name, src)
    finally:
        if os.path.exists(tmp_name):
            os.remove(tmp_name)


def apply_table_width_profiles(hwpx_path, table_headers):
    if not table_headers or not os.path.exists(hwpx_path):
        return
    ns = {'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph'}
    section_name = 'Contents/section0.xml'
    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            section_xml = zf.read(section_name)
    except Exception as e:
        print(f'  [경고] 표 폭 후처리 준비 실패: {e}')
        return
    try:
        ET.register_namespace('hp', ns['hp'])
        root = ET.fromstring(section_xml)
        changed = False
        tables = root.findall('.//hp:tbl', ns)
        for ti, tbl in enumerate(tables):
            if ti >= len(table_headers):
                break
            col_count = int(tbl.attrib.get('colCnt', '0') or 0)
            if col_count <= 1:
                continue
            profile = _table_header_profile(table_headers[ti], col_count)
            if not profile or len(profile) != col_count:
                continue
            total_width = TABLE_TOTAL_WIDTH
            sz = tbl.find('hp:sz', ns)
            if sz is not None:
                total_width = int(sz.attrib.get('width', total_width) or total_width)
            widths = _profile_to_widths(profile, total_width)
            for tc in tbl.findall('.//hp:tc', ns):
                cell_addr = tc.find('hp:cellAddr', ns)
                cell_sz = tc.find('hp:cellSz', ns)
                if cell_addr is None or cell_sz is None:
                    continue
                col = int(cell_addr.attrib.get('colAddr', '0') or 0)
                if 0 <= col < len(widths):
                    cell_sz.set('width', str(widths[col]))
                    changed = True
        if changed:
            _rewrite_zip_entry(hwpx_path, section_name, ET.tostring(root, encoding='utf-8', xml_declaration=True))
    except Exception as e:
        print(f'  [경고] 표 폭 후처리 실패: {e}')


def insert_table(hwp, header, rows):
    all_rows = ([header] if header else []) + rows
    if not all_rows:
        return
    num_rows = len(all_rows)
    num_cols = max(len(r) for r in all_rows)
    col_widths = calc_col_widths(header or [], rows)
    row_heights = calc_row_heights(header or [], rows, col_widths)
    act = hwp.CreateAction('TableCreate')
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem('Rows', num_rows)
    pset.SetItem('Cols', num_cols)
    pset.SetItem('WidthType', 0)
    pset.SetItem('HeightType', 0)
    pset.SetItem('AutoHeight', True)
    for key, value in (('WidthValue', sum(col_widths)), ('HeightValue', sum(row_heights))):
        try:
            pset.SetItem(key, value)
        except Exception:
            pass
    act.Execute(pset)
    moved_right = 0
    try:
        for ci, w in enumerate(col_widths):
            sel_act = hwp.CreateAction('TableColWidth')
            if sel_act is None:
                raise RuntimeError('TableColWidth action unavailable')
            sel_pset = sel_act.CreateSet()
            sel_act.GetDefault(sel_pset)
            sel_pset.SetItem('Width', w)
            sel_act.Execute(sel_pset)
            if ci < num_cols - 1:
                hwp.HAction.Run('TableRightCell')
                moved_right += 1
    except Exception as e:
        print(f'[경고] 열 너비 조정 실패: {e}')
    finally:
        for _ in range(moved_right):
            hwp.HAction.Run('TableLeftCell')
    first_cell = True
    for ri, row in enumerate(all_rows):
        is_header = (ri == 0 and header is not None)
        for ci in range(num_cols):
            if not first_cell:
                hwp.HAction.Run('TableRightCell')
            first_cell = False
            cell_text = row[ci] if ci < len(row) else ''
            if is_header:
                set_para_shape(hwp, align=3)
                set_char_shape(hwp, height=1200, bold=True, font='table')
            else:
                set_para_shape(hwp, align=1)
                set_char_shape(hwp, height=1200, font='table')
            if cell_text:
                insert_text(hwp, cell_text)
    hwp.HAction.Run('MoveDocEnd')
    break_para(hwp)


# ─── 문서 빌드 ─────────────────────────────────────────────────────────────────

def build_doc(hwp, blocks):
    for blk in blocks:
        t = blk.get('type')

        if t == 'h':
            lv = blk['level']
            heights = {1: 1600, 2: 1400, 3: 1300}
            sbefore = {1: 500,  2: 400,  3: 300}
            safter  = {1: 250,  2: 200,  3: 150}
            set_para_shape(hwp, align=1,
                           space_before=sbefore.get(lv, 300),
                           space_after=safter.get(lv, 150))
            set_char_shape(hwp, height=heights.get(lv, 1300), bold=True, font='body')
            insert_text(hwp, blk['text'])
            break_para(hwp)
            set_para_shape(hwp, align=0)
            set_char_shape(hwp, height=1300, font='body')

        elif t == 'p':
            set_para_shape(hwp, align=0)
            set_char_shape(hwp, height=1300, font='body')
            insert_text(hwp, blk['text'])
            break_para(hwp)

        elif t == 'li':
            depth = blk.get('depth', 0)
            set_para_shape(hwp, align=1, indent_left=depth * 400, indent_first=0)
            set_char_shape(hwp, height=1300, font='body')
            insert_text(hwp, blk['text'])
            break_para(hwp)

        elif t == 'bq':
            set_para_shape(hwp, align=1, indent_left=600)
            set_char_shape(hwp, height=1200, italic=True, font='body')
            insert_text(hwp, blk['text'])
            break_para(hwp)

        elif t == 'code':
            set_para_shape(hwp, align=1, indent_left=600)
            set_char_shape(hwp, height=1100, font='table')
            insert_text(hwp, blk['text'])
            break_para(hwp)

        elif t == 'hr':
            set_para_shape(hwp, align=3)
            set_char_shape(hwp, height=1000, font='body')
            insert_text(hwp, '─' * 30)
            break_para(hwp)

        elif t == 'table':
            set_para_shape(hwp, align=0)
            set_char_shape(hwp, height=1200, font='table')
            insert_table(hwp, blk.get('header'), blk.get('rows', []))

        elif t == 'official_header':
            set_para_shape(hwp, align=1)
            set_char_shape(hwp, height=1200, font='table')
            label = blk['key'].ljust(4)
            insert_text(hwp, label + '  ' + blk['value'])
            break_para(hwp)


# ─── 변환 실행 ─────────────────────────────────────────────────────────────────

def build_output_path(src_path, output_dir):
    base_name = os.path.splitext(os.path.basename(src_path))[0]
    out_dir = os.path.abspath(output_dir) if output_dir else os.path.dirname(os.path.abspath(src_path))
    os.makedirs(out_dir, exist_ok=True)
    candidate = os.path.join(out_dir, base_name + '.hwpx')
    if not os.path.exists(candidate):
        return candidate
    for idx in range(2, 1000):
        candidate = os.path.join(out_dir, f'{base_name} - {idx}.hwpx')
        if not os.path.exists(candidate):
            return candidate
    raise FileExistsError(f'저장 가능한 파일명을 찾지 못함: {os.path.join(out_dir, base_name + ".hwpx")}')


def convert_file(hwp, src_path, hwpx_path):
    blocks = detect_and_parse(src_path)
    table_headers = [blk.get('header') or [] for blk in blocks if blk.get('type') == 'table']

    hwp.XHwpDocuments.Add(isTab=False)
    time.sleep(0.5)
    doc = hwp.XHwpDocuments.Item(hwp.XHwpDocuments.Count - 1)

    try:
        build_doc(hwp, blocks)
    except Exception as e:
        print(f'  [경고] 빌드 중 오류: {e}')

    hwp.SaveAs(hwpx_path, 'HWPX', '')
    time.sleep(0.5)
    doc.Close(isDirty=False)
    apply_table_width_profiles(hwpx_path, table_headers)
    time.sleep(0.3)
    ext = os.path.splitext(src_path)[1].upper().lstrip('.')
    print(f'[완료] {ext} → {os.path.basename(hwpx_path)}')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description='Markdown / DOCX → HWPX 변환 (HWP COM 방식)\n'
                    '확장자(.md / .docx)를 자동 감지합니다.',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        'files', nargs='+',
        help='변환할 파일 경로 (.md 또는 .docx, 복수 지정 가능)'
    )
    parser.add_argument(
        '-o', '--output-dir',
        default=None,
        help='저장할 폴더 경로 (기본: 입력 파일과 같은 폴더)'
    )
    args = parser.parse_args()

    print('HWP 실행 중...')
    hwp = win32com.client.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule')
    hwp.XHwpWindows.Item(0).Visible = True
    time.sleep(1.5)

    try:
        for src_path in args.files:
            src_path = os.path.abspath(src_path)
            hwpx_path = build_output_path(src_path, args.output_dir)
            print(f'변환 중: {os.path.basename(src_path)} → {os.path.basename(hwpx_path)}')
            convert_file(hwp, src_path, hwpx_path)
    finally:
        hwp.Quit()

    print('\n전체 변환 완료.')
