"""
HWP COM 자동화로 Markdown → HWPX 변환 (v3)
변경 이력:
  v3 - frontmatter skip / Item(0) 참조 개선 / -o 출력경로 인자 추가
     - 표 열 너비: 한글=2·영문=1 시각 너비 기반 비례 배분 / 전체 너비 14000으로 확대
"""
import win32com.client
import re
import os
import time

# ─── 마크다운 파서 ─────────────────────────────────────────────────────────────

def clean_inline(text):
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

def is_separator(line):
    if len(line) > 500:
        return False
    cells = line.strip().strip('|').split('|')
    return len(cells) >= 1 and all(re.match(r'^[ \t]*:?-+:?[ \t]*$', c) for c in cells)

def parse_table_row(line):
    line = line.strip().strip('|')
    return [clean_inline(c.strip()) for c in line.split('|')]

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
            content = clean_inline(m.group(2))
            return (depth, f'{marker} {content}')
    m = re.match(r'^[-*]\s+(.*)', stripped)
    if m:
        return (0, '• ' + clean_inline(m.group(1)))
    return None


def parse_markdown(text):
    lines = text.splitlines()
    blocks = []
    i = 0
    in_front = False

    while i < len(lines):
        line = lines[i]

        # 빈 줄
        if not line.strip():
            i += 1
            continue

        # frontmatter — 내용은 본문에 출력하지 않고 skip
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
            continue  # frontmatter 내용 무시

        # 공문 구조 헤더 (수신/경유/제목) — 다른 패턴보다 먼저 검사
        stripped_line = line.strip()
        if re.match(r'^(수신|경유|제목)\s*:', stripped_line):
            colon_idx = stripped_line.index(':')
            key = stripped_line[:colon_idx].strip()
            value = clean_inline(stripped_line[colon_idx + 1:].strip())
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
            blocks.append({'type': 'h', 'level': len(m.group(1)), 'text': clean_inline(m.group(2))})
            i += 1
            continue

        # 표
        if line.strip().startswith('|') and i + 1 < len(lines) and is_separator(lines[i + 1]):
            header = parse_table_row(line)
            i += 2
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append(parse_table_row(lines[i]))
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
                blocks.append({'type': 'bq', 'text': clean_inline(text)})
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
        t = clean_inline(line)
        if t:
            blocks.append({'type': 'p', 'text': t})
        i += 1

    return blocks


# ─── HWP COM 헬퍼 ─────────────────────────────────────────────────────────────

def insert_text(hwp, text):
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)

def break_para(hwp):
    hwp.HAction.Run('BreakPara')

def set_char_shape(hwp, height=1300, bold=False, italic=False, font='body'):
    """
    height: 1/100pt 단위 (1300 = 13pt)
    font: 'body' → 휴먼명조 / 'table' → 맑은 고딕
    FaceNameLatin은 영문 가독성을 위해 Arial 유지.
    """
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

def set_para_shape(hwp, align=0, space_before=0, space_after=0, indent_left=0, indent_first=0):
    """align: 0=양쪽 1=왼쪽 2=오른쪽 3=가운데"""
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
    """
    한글·전각문자 = 2, 그 외 = 1로 계산한 시각적 문자 너비.
    빈 문자열은 최소 1 반환.
    """
    w = 0
    for ch in text:
        cp = ord(ch)
        if (0xAC00 <= cp <= 0xD7A3   # 한글 완성형
                or 0x1100 <= cp <= 0x11FF  # 한글 자모
                or 0xA960 <= cp <= 0xA97F  # 한글 자모 확장-A
                or 0xD7B0 <= cp <= 0xD7FF  # 한글 자모 확장-B
                or 0x4E00 <= cp <= 0x9FFF  # CJK 통합 한자
                or 0xFF00 <= cp <= 0xFFEF  # 전각 기호
                or 0x3000 <= cp <= 0x303F):  # CJK 기호·구두점
            w += 2
        else:
            w += 1
    return max(w, 1)


def calc_col_widths(header, rows, total=14000):
    """
    헤더 + 모든 데이터 행의 시각적 너비(한글=2·영문=1) 기반 열 너비 비례 배분.

    total 기본값 14000 ≈ 140mm (A4 본문 영역 150mm에서 좌우 여백 5mm씩 뺀 값).
    각 열 최소 너비: 1500 (15mm).
    단일 열이 전체의 60% 이상 차지하지 않도록 상한 조정.
    """
    all_rows = ([header] if header else []) + (rows if rows else [])
    n = len(header) if header else (len(all_rows[0]) if all_rows else 0)
    if n == 0:
        return []
    if n == 1:
        return [total]

    # 열별 시각적 최대 너비 계산 (헤더 포함, 상한 50자 등가)
    col_vis = []
    for ci in range(n):
        max_w = 1
        for row in all_rows:
            if ci < len(row):
                max_w = max(max_w, _visual_width(row[ci]))
        col_vis.append(min(max_w, 50))

    # 단일 열이 지나치게 넓어지지 않도록 60% 상한 적용
    cap = int(total * 0.6)
    col_vis = [min(w, cap) for w in col_vis]

    total_vis = sum(col_vis) or 1
    result = [max(1500, int(total * w / total_vis)) for w in col_vis]

    # 합계 보정: 최대 열에서 차감/추가
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

    # 첫 행에서 열 너비 순서대로 설정 후 첫 셀로 복귀
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

    # 셀 내용 입력 (행 방향으로 순회)
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
            indent_left = depth * 400
            set_para_shape(hwp, align=1, indent_left=indent_left, indent_first=0)
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
    """
    문서 마지막에 '끝' 표시 삽입.
    - 마지막 블록이 텍스트: '  끝' (2칸 띄움)
    - 마지막 블록이 table: ' 끝' (표 아래 새 단락)
    - 이미 '끝'으로 끝나는 경우 중복 삽입 안 함
    """
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

def convert_file(hwp, md_path, hwpx_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()
    blocks = parse_markdown(md_text)

    # 새 문서를 독립 창으로 열고 참조 보관
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
    print(f'[완료] {os.path.basename(hwpx_path)}')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description='Markdown → HWPX 변환 (HWP COM 방식)'
    )
    parser.add_argument('files', nargs='+', help='변환할 .md 파일 경로 (복수 지정 가능)')
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
        for md_path in args.files:
            md_path = os.path.abspath(md_path)
            base_name = os.path.splitext(os.path.basename(md_path))[0]
            out_dir = os.path.abspath(args.output_dir) if args.output_dir else os.path.dirname(md_path)
            os.makedirs(out_dir, exist_ok=True)
            hwpx_path = os.path.join(out_dir, base_name + '.hwpx')
            print(f'변환 중: {os.path.basename(md_path)} → {hwpx_path}')
            convert_file(hwp, md_path, hwpx_path)
    finally:
        hwp.Quit()

    print('\n전체 변환 완료.')
