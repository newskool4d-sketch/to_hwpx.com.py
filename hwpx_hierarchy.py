import re
from pathlib import Path

from document_hierarchy import hwpx_para_metrics, hwpx_para_pr_id, parse_hierarchy_item


def _para_pr_xml(para_pr_id: str, left: int, intent: int) -> str:
    return (
        f'<hh:paraPr id="{para_pr_id}" tabPrIDRef="0" condense="0" fontLineHeight="0" '
        f'snapToGrid="1" suppressLineNumbers="0" checked="0">'
        f'<hh:align horizontal="JUSTIFY" vertical="BASELINE"/>'
        f'<hh:heading type="NONE" idRef="0" level="0"/>'
        f'<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" '
        f'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        f'<hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
        f'<hh:margin><hc:intent value="{intent}" unit="HWPUNIT"/><hc:left value="{left}" unit="HWPUNIT"/>'
        f'<hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/>'
        f'<hc:next value="0" unit="HWPUNIT"/></hh:margin>'
        f'<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/></hp:case>'
        f'<hp:default><hh:margin><hc:intent value="{intent}" unit="HWPUNIT"/>'
        f'<hc:left value="{left}" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        f'<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin>'
        f'<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/></hp:default></hp:switch>'
        f'<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" offsetTop="0" '
        f'offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>'
    )


def write_header_with_hierarchy(src_header: Path, dst_header: Path) -> None:
    text = src_header.read_text(encoding='utf-8')
    if 'id="200"' in text:
        dst_header.write_text(text, encoding='utf-8')
        return
    para_props = re.search(r'<hh:paraProperties itemCnt="(\d+)">', text)
    if para_props is None:
        raise RuntimeError('header.xml에서 hh:paraProperties를 찾지 못함')
    item_count = int(para_props.group(1))
    additions = []
    for depth in range(9):
        left, intent = hwpx_para_metrics(depth)
        additions.append(_para_pr_xml(hwpx_para_pr_id(depth), left, intent))
    updated = text.replace(
        para_props.group(0),
        f'<hh:paraProperties itemCnt="{item_count + len(additions)}">',
        1,
    )
    updated = updated.replace('</hh:paraProperties>', ''.join(additions) + '</hh:paraProperties>', 1)
    dst_header.write_text(updated, encoding='utf-8')


def make_hierarchy_para(
    helper,
    text: str,
    fallback_depth: int = 0,
    marker: str | None = None,
    content: str | None = None,
) -> str:
    if marker is not None and content is not None:
        depth = fallback_depth
        item_marker = marker
        item_content = content
    else:
        item = parse_hierarchy_item(text)
        if item is None:
            depth = fallback_depth
            item_marker = '-'
            item_content = text
        else:
            depth = item.depth
            item_marker = item.marker
            item_content = item.content
    p_id = helper.next_id()
    para_pr = hwpx_para_pr_id(depth)
    return (
        f'<hp:p id="{p_id}" paraPrIDRef="{para_pr}" styleIDRef="0" pageBreak="0" '
        f'columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="18"><hp:t>{helper.xml_escape(item_marker + " ")}</hp:t></hp:run>'
        f'<hp:run charPrIDRef="38"><hp:t>{helper.xml_escape(item_content)}</hp:t></hp:run></hp:p>'
    )
