"""
HWP COM 자동화로 Markdown / DOCX → HWPX 변환 (v1)
확장자(.md / .docx)를 자동 감지하여 적절한 파서로 처리.
이미지는 skip.

변경 이력:
  v1 - md_to_hwpx_com v3 + docx_to_hwpx_com v1 통합
"""
import win32com.client
import re
import os
import time


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

def _visual_width(text):
    w = 0
    for ch in text:
        cp = ord(ch)
        if (0xAC00 <= cp <= 0xD7A3
                or 0x1100 <= cp <= 0x11FF
                or 0xA960 <= cp <= 0xA97F
                or 0xD7B0 <= cp <= 0xD7FF
                or 0x4E00 <= cp <= 0x9FFF
                or 0xFF00 <= cp <= 0xFFEF
                or 0x3000 <= cp <= 0x303F):
            w += 2
        else:
            w += 1
    return max(w, 1)


def calc_col_widths(header, rows, total=14000):
    all_rows = ([header] if header else []) + (rows if rows else [])
    n = len(header) if header else (len(all_rows[0]) if all_rows else 0)
    if n == 0:
        return []
    if n == 1:
        return [total]
    col_vis = []
    for ci in range(n):
        max_w = 1
        for row in all_rows:
            if ci < len(row):
                max_w = max(max_w, _visual_width(row[ci]))
        col_vis.append(min(max_w, 50))
    cap = int(total * 0.6)
    col_vis = [min(w, cap) for w in col_vis]
    total_vis = sum(col_vis) or 1
    result = [max(1500, int(total * w / total_vis)) for w in col_vis]
    diff = total - sum(result)
    if diff != 0:
        max_idx = result.index(max(result))
        result[max_idx] += diff
    return result


def insert_table(hwp, header, rows):
    all_rows = ([header] if header else []) + rows
    if not all_rows:
        return
    num_rows = len(all_rows)
    num_cols = max(len(r) for r in all_rows)
    col_widths = calc_col_widths(header or [], rows)
    act = hwp.CreateAction('TableCreate')
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem('Rows', num_rows)
    pset.SetItem('Cols', num_cols)
    pset.SetItem('WidthType', 0)
    pset.SetItem('HeightType', 0)
    pset.SetItem('AutoHeight', True)
    act.Execute(pset)
    try:
        for ci, w in enumerate(col_widths):
            sel_act = hwp.CreateAction('TableColWidth')
            sel_pset = sel_act.CreateSet()
            sel_act.GetDefault(sel_pset)
            sel_pset.SetItem('Width', w)
            sel_act.Execute(sel_pset)
            if ci < num_cols - 1:
                hwp.HAction.Run('TableRightCell')
        for _ in range(num_cols - 1):
            hwp.HAction.Run('TableLeftCell')
    except Exception as e:
        print(f'[경고] 열 너비 조정 실패: {e}')
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


def _insert_end_mark(hwp, blocks):
    if not blocks:
        return
    last = blocks[-1]
    last_text = last.get('text', '') or ''
    if last_text.strip().endswith('끝'):
        return
    if last['type'] == 'table':
        last_rows = last.get('rows', [])
        if last_rows:
            last_row_text = ' '.join(last_rows[-1])
            if last_row_text.strip().endswith('끝') or last_row_text.strip() == '이하 빈칸':
                return
        hwp.HAction.Run('MoveDocEnd')
        set_para_shape(hwp, align=1)
        set_char_shape(hwp, height=1300, font='body')
        insert_text(hwp, ' 끝')
        break_para(hwp)
    else:
        hwp.HAction.Run('MoveDocEnd')
        set_para_shape(hwp, align=1)
        set_char_shape(hwp, height=1300, font='body')
        insert_text(hwp, '  끝')
        break_para(hwp)


# ─── 변환 실행 ─────────────────────────────────────────────────────────────────

def convert_file(hwp, src_path, hwpx_path):
    blocks = detect_and_parse(src_path)

    hwp.XHwpDocuments.Add(isTab=False)
    time.sleep(0.5)
    doc = hwp.XHwpDocuments.Item(hwp.XHwpDocuments.Count - 1)

    try:
        build_doc(hwp, blocks)
        _insert_end_mark(hwp, blocks)
    except Exception as e:
        print(f'  [경고] 빌드 중 오류: {e}')

    hwp.SaveAs(hwpx_path, 'HWPX', '')
    time.sleep(0.5)
    doc.Close(isDirty=False)
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
            base_name = os.path.splitext(os.path.basename(src_path))[0]
            out_dir = os.path.abspath(args.output_dir) if args.output_dir else os.path.dirname(src_path)
            os.makedirs(out_dir, exist_ok=True)
            hwpx_path = os.path.join(out_dir, base_name + '.hwpx')
            print(f'변환 중: {os.path.basename(src_path)} → {os.path.basename(hwpx_path)}')
            convert_file(hwp, src_path, hwpx_path)
    finally:
        hwp.Quit()

    print('\n전체 변환 완료.')
