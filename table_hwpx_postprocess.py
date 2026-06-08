import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from table_settings import TABLE_TOTAL_WIDTH, exact_header_widths


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
    except (KeyError, OSError, zipfile.BadZipFile) as e:
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
            total_width = TABLE_TOTAL_WIDTH
            sz = tbl.find('hp:sz', ns)
            if sz is not None:
                total_width = int(sz.attrib.get('width', total_width) or total_width)
            widths = exact_header_widths(table_headers[ti], col_count, total_width)
            if not widths:
                continue
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
    except (ET.ParseError, OSError, ValueError, zipfile.BadZipFile) as e:
        print(f'  [경고] 표 폭 후처리 실패: {e}')
