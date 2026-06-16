import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from table_settings import TABLE_TOTAL_WIDTH, exact_header_widths
from table_layout import TableLayout
from table_layout import TableCellStyle
from table_layout import table_layout_for
from table_hwpx_styles import ensure_table_border_fills
from table_hwpx_styles import register_hwpx_namespaces


HP_NS = 'http://www.hancom.co.kr/hwpml/2011/paragraph'


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


def _scaled_widths(widths, total_width):
    width_sum = sum(widths)
    if not widths or width_sum <= 0:
        return []
    if width_sum == total_width:
        return list(widths)
    scaled = [max(1, int(total_width * width / width_sum)) for width in widths]
    diff = total_width - sum(scaled)
    if diff:
        target = max(range(len(widths)), key=lambda index: widths[index])
        scaled[target] += diff
    return scaled


def _layout_widths(layout, col_count, total_width):
    if isinstance(layout, TableLayout):
        if len(layout.column_widths) == col_count:
            return _scaled_widths(layout.column_widths, total_width)
        return []
    if isinstance(layout, list):
        exact_widths = exact_header_widths(layout, col_count, total_width)
        if exact_widths:
            return exact_widths
        return table_layout_for(layout, [], total_width=total_width).column_widths
    return []


def _layout_style(layout):
    if isinstance(layout, TableLayout):
        return layout.style
    return TableCellStyle()


def _ensure_cell_margin(tc, ns, style):
    cell_margin = tc.find('hp:cellMargin', ns)
    if cell_margin is None:
        cell_margin = ET.SubElement(tc, f'{{{HP_NS}}}cellMargin')
    cell_margin.set('left', str(style.margin_left))
    cell_margin.set('right', str(style.margin_right))
    cell_margin.set('top', str(style.margin_top))
    cell_margin.set('bottom', str(style.margin_bottom))


def _span_by_addr(layout):
    if not isinstance(layout, TableLayout):
        return {}
    return {(span[0], span[1]): span for span in layout.merged_cells if len(span) == 4}


def _ensure_cell_span(tc, ns):
    cell_span = tc.find('hp:cellSpan', ns)
    if cell_span is None:
        cell_span = ET.SubElement(tc, f'{{{HP_NS}}}cellSpan')
    return cell_span


def apply_table_width_profiles(hwpx_path, table_layouts):
    if not table_layouts or not os.path.exists(hwpx_path):
        return
    ns = {'hp': HP_NS}
    section_name = 'Contents/section0.xml'
    header_name = 'Contents/header.xml'
    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            section_xml = zf.read(section_name)
            header_xml = zf.read(header_name)
    except (KeyError, OSError, zipfile.BadZipFile) as e:
        print(f'  [경고] 표 폭 후처리 준비 실패: {e}', file=sys.stderr)
        return
    try:
        register_hwpx_namespaces()
        root = ET.fromstring(section_xml)
        header_root = ET.fromstring(header_xml)
        changed = False
        header_changed = False
        tables = root.findall('.//hp:tbl', ns)
        for ti, tbl in enumerate(tables):
            if ti >= len(table_layouts):
                break
            layout = table_layouts[ti]
            col_count = int(tbl.attrib.get('colCnt', '0') or 0)
            if col_count <= 1:
                continue
            total_width = TABLE_TOTAL_WIDTH
            sz = tbl.find('hp:sz', ns)
            if sz is not None:
                total_width = int(sz.attrib.get('width', total_width) or total_width)
            widths = _layout_widths(layout, col_count, total_width)
            if not widths:
                continue
            style = _layout_style(layout)
            border_refs = ensure_table_border_fills(header_root, style)
            header_changed = header_changed or border_refs.changed
            tbl.set('borderFillIDRef', border_refs.body_id)
            spans = _span_by_addr(layout)
            for tc in tbl.findall('.//hp:tc', ns):
                cell_addr = tc.find('hp:cellAddr', ns)
                cell_sz = tc.find('hp:cellSz', ns)
                if cell_addr is None or cell_sz is None:
                    continue
                col = int(cell_addr.attrib.get('colAddr', '0') or 0)
                if 0 <= col < len(widths):
                    row = int(cell_addr.attrib.get('rowAddr', '0') or 0)
                    row_span = 1
                    col_span = 1
                    span = spans.get((row, col))
                    if span is not None:
                        row_span = span[2]
                        col_span = span[3]
                        cell_span = _ensure_cell_span(tc, ns)
                        cell_span.set('rowSpan', str(row_span))
                        cell_span.set('colSpan', str(col_span))
                    cell_width = sum(widths[col : min(len(widths), col + col_span)])
                    cell_sz.set('width', str(cell_width))
                    if row == 0:
                        tc.set('header', '1')
                        tc.set('borderFillIDRef', border_refs.header_id)
                    else:
                        tc.set('borderFillIDRef', border_refs.body_id)
                    tc.set('hasMargin', '1')
                    _ensure_cell_margin(tc, ns, style)
                    changed = True
            changed = True
        if header_changed:
            _rewrite_zip_entry(hwpx_path, header_name, ET.tostring(header_root, encoding='utf-8', xml_declaration=True))
        if changed:
            _rewrite_zip_entry(hwpx_path, section_name, ET.tostring(root, encoding='utf-8', xml_declaration=True))
    except (ET.ParseError, OSError, ValueError, zipfile.BadZipFile) as e:
        print(f'  [경고] 표 폭 후처리 실패: {e}', file=sys.stderr)
