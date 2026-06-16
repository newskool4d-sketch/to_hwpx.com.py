import importlib
import sys

from document_hierarchy import hwp_com_style, parse_hierarchy_item
from table_settings import TABLE_TOTAL_WIDTH, calc_row_heights
from table_layout import table_layout_for


class HwpActionUnavailableError(RuntimeError):
    pass


def hwp_runtime_exception_types() -> tuple[type[BaseException], ...]:
    base_types: list[type[BaseException]] = [RuntimeError, AttributeError, OSError, TypeError]
    try:
        pywintypes = importlib.import_module("pywintypes")
    except ImportError:
        return tuple(base_types)
    com_error = getattr(pywintypes, "com_error", None)
    if isinstance(com_error, type) and issubclass(com_error, BaseException):
        base_types.append(com_error)
    return tuple(base_types)


def insert_text(hwp, text):
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)


def break_para(hwp):
    hwp.HAction.Run('BreakPara')


def set_char_shape(hwp, height=1300, bold=False, italic=False, font='body'):
    face_hangul = '휴먼명조' if font == 'body' else '맑은 고딕'
    face_latin = 'Arial'
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
    act = hwp.CreateAction('ParagraphShape')
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem('Align', align)
    pset.SetItem('SpaceBefore', space_before)
    pset.SetItem('SpaceAfter', space_after)
    pset.SetItem('IndentLeft', indent_left)
    pset.SetItem('IndentFirst', indent_first)
    act.Execute(pset)


def _clean_column_widths(column_widths):
    if not isinstance(column_widths, list):
        return None
    if not all(type(value) is int for value in column_widths):
        return None
    return column_widths


def _clean_total_width(total_width):
    if type(total_width) is int and total_width > 0:
        return total_width
    return TABLE_TOTAL_WIDTH


def insert_table(hwp, header, rows, table_role=None, column_widths=None, total_width=None):
    all_rows = ([header] if header else []) + rows
    if not all_rows:
        return
    num_rows = len(all_rows)
    num_cols = max(len(r) for r in all_rows)
    role = table_role if isinstance(table_role, str) else None
    layout = table_layout_for(header or [], rows, role, _clean_column_widths(column_widths), total_width=_clean_total_width(total_width))
    col_widths = layout.column_widths
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
        except hwp_runtime_exception_types() as exc:
            print(f'[경고] 표 크기 설정 실패({key}): {exc}', file=sys.stderr)
    act.Execute(pset)
    moved_right = 0
    try:
        for ci, w in enumerate(col_widths):
            sel_act = hwp.CreateAction('TableColWidth')
            if sel_act is None:
                raise HwpActionUnavailableError('TableColWidth action unavailable')
            sel_pset = sel_act.CreateSet()
            sel_act.GetDefault(sel_pset)
            sel_pset.SetItem('Width', w)
            sel_act.Execute(sel_pset)
            if ci < num_cols - 1:
                hwp.HAction.Run('TableRightCell')
                moved_right += 1
    except hwp_runtime_exception_types() as e:
        print(f'[경고] 열 너비 조정 실패: {e}', file=sys.stderr)
    finally:
        for _ in range(moved_right):
            hwp.HAction.Run('TableLeftCell')
    first_cell = True
    for ri, row in enumerate(all_rows):
        is_header = ri == 0 and header is not None
        for ci in range(num_cols):
            if not first_cell:
                hwp.HAction.Run('TableRightCell')
            first_cell = False
            cell_text = row[ci] if ci < len(row) else ''
            if is_header:
                set_para_shape(hwp, align=layout.style.header_align)
                set_char_shape(hwp, height=1200, bold=True, font='table')
            else:
                set_para_shape(hwp, align=layout.style.body_align)
                set_char_shape(hwp, height=1200, font='table')
            if cell_text:
                insert_text(hwp, cell_text)
    hwp.HAction.Run('MoveDocEnd')
    break_para(hwp)


def build_doc(hwp, blocks, table_total_width=None):
    for blk in blocks:
        t = blk.get('type')

        if t == 'h':
            lv = blk['level']
            heights = {1: 1600, 2: 1400, 3: 1300}
            sbefore = {1: 500, 2: 400, 3: 300}
            safter = {1: 250, 2: 200, 3: 150}
            set_para_shape(hwp, align=1, space_before=sbefore.get(lv, 300), space_after=safter.get(lv, 150))
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
            text = blk['text']
            item = parse_hierarchy_item(text)
            marker = str(blk.get('marker') or (item.marker if item is not None else '-'))
            depth = int(blk.get('depth', item.depth if item is not None else 0) or 0)
            style = hwp_com_style(depth, marker)
            set_para_shape(hwp, align=1, indent_left=style.left, indent_first=style.first)
            set_char_shape(hwp, height=1300, font='body')
            insert_text(hwp, text)
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
            insert_table(
                hwp,
                blk.get('header'),
                blk.get('rows', []),
                table_role=blk.get('table_role'),
                column_widths=blk.get('column_widths'),
                total_width=table_total_width,
            )

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
