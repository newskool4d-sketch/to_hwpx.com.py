import argparse
import importlib.util
import re
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from hwpx_hierarchy import make_hierarchy_para, write_header_with_hierarchy
from table_settings import calc_col_widths, calc_row_heights
from to_hwpx_com import detect_and_parse


SKILL_DIR = Path(r'C:\Users\홍주형\.agents\skills\hwpx변환')
SCRIPTS_DIR = SKILL_DIR / 'scripts'
TABLE_BODY_WIDTH = 42520
TABLE_MIN_RENDER_HEIGHT = 2200

helper_spec = importlib.util.spec_from_file_location('hwpx_helpers', SCRIPTS_DIR / 'hwpx_helpers.py')
if helper_spec is None or helper_spec.loader is None:
    raise ImportError('hwpx_helpers.py를 찾을 수 없습니다.')
hwpx_helpers = importlib.util.module_from_spec(helper_spec)
helper_spec.loader.exec_module(hwpx_helpers)


def _output_path(src_path, output=None):
    if output:
        return Path(output).expanduser().resolve()
    candidate = src_path.with_suffix('.hwpx')
    if not candidate.exists():
        return candidate
    for index in range(2, 1000):
        candidate = src_path.with_name(f'{src_path.stem} - {index}.hwpx')
        if not candidate.exists():
            return candidate
    raise FileExistsError(f'저장 가능한 출력 파일명을 찾지 못함: {src_path}')


def _document_title(src_path, blocks):
    for block in blocks:
        if block.get('type') == 'h' and block.get('level') == 1 and block.get('text'):
            return block['text']
    return src_path.stem


def _split_section_title(text, fallback_number):
    match = re.match(r'^(\d+(?:\.\d+)*)\.\s*(.+)$', text or '')
    if match:
        return match.group(1), match.group(2)
    return str(fallback_number), text or ''


def _normalize_rows(header, rows):
    all_rows = ([header] if header else []) + (rows or [])
    col_count = max((len(row) for row in all_rows), default=0)
    normalized = []
    for row in all_rows:
        values = [str(cell or '').strip() for cell in row]
        normalized.append(values + [''] * (col_count - len(values)))
    return normalized, col_count


def _make_table_cell(text, col_idx, row_idx, width, height, is_header):
    cell_para_id = hwpx_helpers.next_id()
    border_fill = '13' if is_header else '4'
    char_pr = '18' if is_header else '38'
    para_pr = '21' if is_header else '4'
    header_flag = '1' if is_header else '0'
    return (
        f'<hp:tc name="" header="{header_flag}" hasMargin="0" protect="0" editable="0" '
        f'dirty="1" borderFillIDRef="{border_fill}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" '
        f'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" '
        f'hasTextRef="0" hasNumRef="0">'
        f'<hp:p paraPrIDRef="{para_pr}" styleIDRef="0" pageBreak="0" columnBreak="0" '
        f'merged="0" id="{cell_para_id}">'
        f'<hp:run charPrIDRef="{char_pr}"><hp:t>{hwpx_helpers.xml_escape(text)}</hp:t></hp:run>'
        f'</hp:p></hp:subList>'
        f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{width}" height="{height}"/>'
        f'<hp:cellMargin left="170" right="170" top="120" bottom="120"/></hp:tc>'
    )


def make_table(header, rows):
    normalized, col_count = _normalize_rows(header or [], rows or [])
    if not normalized or col_count == 0:
        return ''
    has_header = bool(header)
    widths = calc_col_widths(header or [], rows or [], total=TABLE_BODY_WIDTH)
    if len(widths) < col_count:
        widths.extend([TABLE_BODY_WIDTH // col_count] * (col_count - len(widths)))
    width_diff = TABLE_BODY_WIDTH - sum(widths[:col_count])
    widths[col_count - 1] += width_diff

    heights = calc_row_heights(header or [], rows or [], widths[:col_count])
    if len(heights) < len(normalized):
        heights.extend([TABLE_MIN_RENDER_HEIGHT] * (len(normalized) - len(heights)))
    heights = [max(TABLE_MIN_RENDER_HEIGHT, height) for height in heights]

    p_id = hwpx_helpers.next_id()
    tbl_id = hwpx_helpers.next_id()
    table_rows = []
    for row_idx, row in enumerate(normalized):
        cells = []
        is_header = has_header and row_idx == 0
        for col_idx in range(col_count):
            cells.append(_make_table_cell(row[col_idx], col_idx, row_idx, widths[col_idx], heights[row_idx], is_header))
        table_rows.append(f'<hp:tr>{"".join(cells)}</hp:tr>')

    total_height = sum(heights)
    return (
        f'<hp:p id="{p_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">'
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
        f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" '
        f'rowCnt="{len(normalized)}" colCnt="{col_count}" cellSpacing="0" borderFillIDRef="4" noAdjust="0">'
        f'<hp:sz width="{TABLE_BODY_WIDTH}" widthRelTo="ABSOLUTE" height="{total_height}" '
        f'heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
        f'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
        f'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'{"".join(table_rows)}</hp:tbl></hp:run></hp:p>'
    )


def build_section_xml(src_path, section_path):
    blocks = detect_and_parse(src_path)
    title = _document_title(src_path, blocks)
    hwpx_helpers.reset_id()

    header_path = SKILL_DIR / 'templates' / 'government' / 'header.xml'
    ref_hwpx = SKILL_DIR / 'assets' / 'government-reference.hwpx'
    hwpx_helpers.validate_header_for_government(header_path)
    secpr, colpr = hwpx_helpers.extract_secpr_and_colpr(ref_hwpx)

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
        f'<hs:sec {hwpx_helpers.NS_DECL}>',
        hwpx_helpers.make_first_para(secpr, colpr),
    ]
    today = datetime.now().strftime('%Y. %m. %d.')
    parts.extend(hwpx_helpers.make_cover_page(title, subtitle='Markdown 변환 문서', date=today))

    section_count = 0
    for block in blocks:
        block_type = block.get('type')
        if block_type == 'h':
            level = block.get('level', 1)
            text = block.get('text', '')
            if level == 1 and text == title:
                continue
            if level == 2:
                section_count += 1
                number, section_title = _split_section_title(text, section_count)
                parts.append(hwpx_helpers.make_section_bar(number, section_title))
                parts.append(hwpx_helpers.make_empty_line())
            else:
                parts.append(hwpx_helpers.make_text_para(text, charpr='18', parapr='4'))
        elif block_type == 'table':
            table_xml = make_table(block.get('header') or [], block.get('rows') or [])
            if table_xml:
                parts.append(table_xml)
                parts.append(hwpx_helpers.make_empty_line())
        elif block_type == 'li':
            depth = int(block.get('depth', 0) or 0)
            parts.append(
                make_hierarchy_para(
                    hwpx_helpers,
                    block.get('text', ''),
                    fallback_depth=depth,
                    marker=block.get('marker'),
                    content=block.get('content'),
                )
            )
        elif block_type == 'bq':
            parts.append(hwpx_helpers.make_body_para('※', block.get('text', '')))
        elif block_type == 'hr':
            parts.append(hwpx_helpers.make_empty_line())
        elif block_type == 'official_header':
            parts.append(hwpx_helpers.make_body_para(block.get('key', ''), block.get('value', '')))
        else:
            text = block.get('text', '')
            if text:
                parts.append(hwpx_helpers.make_text_para(text, charpr='38', parapr='4'))
    parts.append('</hs:sec>')
    section_path.write_text('\n'.join(parts), encoding='utf-8')
    return title


def convert_markdown(src_path, output=None):
    src_path = Path(src_path).expanduser().resolve()
    if not src_path.is_file():
        raise FileNotFoundError(src_path)
    output_path = _output_path(src_path, output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix='md_hwpx_') as tmpdir:
        header_path = Path(tmpdir) / 'header.xml'
        section_path = Path(tmpdir) / 'section0.xml'
        write_header_with_hierarchy(SKILL_DIR / 'templates' / 'government' / 'header.xml', header_path)
        title = build_section_xml(src_path, section_path)
        subprocess.run(
            [
                sys.executable,
                str(SCRIPTS_DIR / 'build_hwpx.py'),
                '--header',
                str(header_path),
                '--section',
                str(section_path),
                '--title',
                title,
                '--creator',
                'Codex',
                '--output',
                str(output_path),
            ],
            check=True,
        )
        subprocess.run([sys.executable, str(SCRIPTS_DIR / 'fix_namespaces.py'), str(output_path)], check=True)
        subprocess.run([sys.executable, str(SCRIPTS_DIR / 'validate.py'), str(output_path)], check=True)
    return output_path


def main(argv=None):
    parser = argparse.ArgumentParser(description='Markdown 파일을 COM 없이 HWPX로 직접 변환')
    parser.add_argument('markdown_path', help='변환할 Markdown 파일')
    parser.add_argument('-o', '--output', default=None, help='출력 HWPX 경로')
    args = parser.parse_args(argv)
    output_path = convert_markdown(args.markdown_path, output=args.output)
    print(output_path)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
