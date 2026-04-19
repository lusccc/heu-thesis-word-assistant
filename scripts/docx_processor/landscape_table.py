"""横向页面表格处理模块

解析thesis.qmd中的 <!-- landscape: caption关键词 --> 注释，
在编译后的docx中将对应表格所在页面设为横向（landscape）。
"""
import copy
import xml.etree.ElementTree as ET
import re
import os
from .utils import NAMESPACES
from .config import DEFAULT_QMD_FILENAME

# A4默认页面参数（缇）
PAGE_W_PORTRAIT = '11906'
PAGE_H_PORTRAIT = '16838'


def load_landscape_captions_from_qmd(qmd_path=None):
    """从thesis.qmd读取需要横向显示的表格caption关键词

    格式: <!-- landscape: caption关键词 -->
    关键词用于在编译后的docx中匹配表格标题段落。

    Returns:
        list[str]: caption关键词列表
    """
    if qmd_path is None:
        qmd_path = DEFAULT_QMD_FILENAME

    captions = []
    if not os.path.exists(qmd_path):
        return captions

    with open(qmd_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        match = re.match(r'^\s*<!--\s*landscape\s*:\s*(.+?)\s*-->', line)
        if match:
            caption = match.group(1).strip()
            captions.append(caption)

    return captions


def _get_page_settings_from_body(root, w_ns):
    """从文档body/sectPr中提取当前页面设置"""
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return None

    # body的最后一个子元素通常是sectPr
    sect_pr = None
    for child in reversed(list(body)):
        if child.tag == f'{{{w_ns}}}sectPr':
            sect_pr = child
            break

    if sect_pr is None:
        return None

    settings = {}
    pgSz = sect_pr.find(f'{{{w_ns}}}pgSz')
    if pgSz is not None:
        settings['w'] = pgSz.get(f'{{{w_ns}}}w', PAGE_W_PORTRAIT)
        settings['h'] = pgSz.get(f'{{{w_ns}}}h', PAGE_H_PORTRAIT)
    else:
        settings['w'] = PAGE_W_PORTRAIT
        settings['h'] = PAGE_H_PORTRAIT

    pgMar = sect_pr.find(f'{{{w_ns}}}pgMar')
    if pgMar is not None:
        for attr in ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter']:
            settings[f'margin_{attr}'] = pgMar.get(f'{{{w_ns}}}{attr}', '0')
    else:
        settings.update({
            'margin_top': '1587', 'margin_bottom': '1587',
            'margin_left': '1417', 'margin_right': '1417',
            'margin_header': '1134', 'margin_footer': '1134',
            'margin_gutter': '0',
        })

    return settings


def _find_nearest_sectpr(body, ref_idx, w_ns):
    """从ref_idx向前查找最近的已有sectPr元素（含headerRef等完整属性）"""
    body_children = list(body)
    # 先向前查找段落中的sectPr
    for k in range(ref_idx - 1, -1, -1):
        elem = body_children[k]
        if elem.tag == f'{{{w_ns}}}p':
            sect = elem.find(f'.//{{{w_ns}}}sectPr')
            if sect is not None:
                return sect
    # 如果都没有，使用body末尾的sectPr
    for child in reversed(body_children):
        if child.tag == f'{{{w_ns}}}sectPr':
            return child
    return None


def _create_section_break_para(w_ns, ref_sect_pr, page_settings, landscape=False):
    """创建包含sectPr的空段落（分节符）

    从参考sectPr深拷贝完整属性（headerReference、footerReference、
    cols、docGrid等），只修改pgSz的尺寸和orient。

    Args:
        w_ns: Word命名空间URI
        ref_sect_pr: 参考sectPr元素（用于克隆headerRef等属性）
        page_settings: 页面设置字典
        landscape: True则创建横向分节符，False则创建纵向分节符

    Returns:
        ET.Element: 包含sectPr的段落元素
    """
    p = ET.Element(f'{{{w_ns}}}p')
    pPr = ET.SubElement(p, f'{{{w_ns}}}pPr')

    if ref_sect_pr is not None:
        sectPr = copy.deepcopy(ref_sect_pr)
        # 清除rsidR/rsidSect等修订标记（避免冲突）
        for attr_name in list(sectPr.attrib.keys()):
            if 'rsid' in attr_name.lower():
                del sectPr.attrib[attr_name]
        # 移除pgNumType（避免页码重置）
        for pn in list(sectPr.findall(f'{{{w_ns}}}pgNumType')):
            sectPr.remove(pn)
    else:
        sectPr = ET.Element(f'{{{w_ns}}}sectPr')

    pPr.append(sectPr)

    # 确保type为nextPage
    typ = sectPr.find(f'{{{w_ns}}}type')
    if typ is None:
        typ = ET.SubElement(sectPr, f'{{{w_ns}}}type')
    typ.set(f'{{{w_ns}}}val', 'nextPage')

    # 修改pgSz
    pgSz = sectPr.find(f'{{{w_ns}}}pgSz')
    if pgSz is None:
        pgSz = ET.SubElement(sectPr, f'{{{w_ns}}}pgSz')

    orient_key = f'{{{w_ns}}}orient'
    if landscape:
        pgSz.set(f'{{{w_ns}}}w', page_settings['h'])
        pgSz.set(f'{{{w_ns}}}h', page_settings['w'])
        pgSz.set(orient_key, 'landscape')
    else:
        pgSz.set(f'{{{w_ns}}}w', page_settings['w'])
        pgSz.set(f'{{{w_ns}}}h', page_settings['h'])
        if orient_key in pgSz.attrib:
            del pgSz.attrib[orient_key]

    # 确保pgMar存在
    pgMar = sectPr.find(f'{{{w_ns}}}pgMar')
    if pgMar is None:
        pgMar = ET.SubElement(sectPr, f'{{{w_ns}}}pgMar')
        for attr in ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter']:
            pgMar.set(f'{{{w_ns}}}{attr}', page_settings[f'margin_{attr}'])

    return p


def _get_paragraph_text(para, w_ns):
    """获取段落的纯文本"""
    text = ''
    for t in para.iter(f'{{{w_ns}}}t'):
        if t.text:
            text += t.text
    return text


def _scale_table_to_page_width(tbl, target_width, w_ns):
    """按比例缩放表格列宽以匹配目标页面宽度

    调整 tblGrid/gridCol 和所有单元格的 tcW 值。

    Args:
        tbl: 表格元素 (w:tbl)
        target_width: 目标可用宽度（缇）
        w_ns: Word命名空间URI

    Returns:
        bool: 是否成功缩放
    """
    tbl_grid = tbl.find(f'{{{w_ns}}}tblGrid')
    if tbl_grid is None:
        return False

    grid_cols = tbl_grid.findall(f'{{{w_ns}}}gridCol')
    if not grid_cols:
        return False

    # 计算当前gridCol总宽
    current_total = 0
    col_widths = []
    for gc in grid_cols:
        w_val = gc.get(f'{{{w_ns}}}w', '0')
        width = int(w_val)
        col_widths.append(width)
        current_total += width

    if current_total <= 0:
        return False

    scale = target_width / current_total
    if abs(scale - 1.0) < 0.05:
        return False  # 无需缩放

    # 按比例调整gridCol
    new_total = 0
    for i, gc in enumerate(grid_cols):
        new_w = round(col_widths[i] * scale)
        gc.set(f'{{{w_ns}}}w', str(new_w))
        new_total += new_w

    # 同时调整所有行中单元格的tcW（如果存在）
    for tr in tbl.iter(f'{{{w_ns}}}tr'):
        for tc in tr.findall(f'{{{w_ns}}}tc'):
            tcPr = tc.find(f'{{{w_ns}}}tcPr')
            if tcPr is None:
                continue
            tcW = tcPr.find(f'{{{w_ns}}}tcW')
            if tcW is None:
                continue
            w_type = tcW.get(f'{{{w_ns}}}type', '')
            if w_type == 'dxa':
                old_w = int(tcW.get(f'{{{w_ns}}}w', '0'))
                tcW.set(f'{{{w_ns}}}w', str(round(old_w * scale)))

    print(f"    列宽缩放: {current_total} → {new_total} twips (×{scale:.2f})")
    return True


def process_landscape_tables(root, qmd_path=None):
    """将标记为landscape的表格所在页面设为横向

    通过标题文本匹配定位表格，在表格（含标题）前后插入分节符：
    - 表格标题前：portrait sectPr（结束前一个纵向节）
    - 表格（或表注）后：landscape sectPr（结束横向节）

    Args:
        root: 文档根元素
        qmd_path: thesis.qmd路径

    Returns:
        int: 处理的表格数量
    """
    landscape_captions = load_landscape_captions_from_qmd(qmd_path)
    if not landscape_captions:
        return 0

    print(f"正在处理横向表格（共 {len(landscape_captions)} 个）...")
    for cap in landscape_captions:
        print(f"    关键词: {cap}")

    w_ns = NAMESPACES['w']
    page_settings = _get_page_settings_from_body(root, w_ns)
    if page_settings is None:
        print("  警告：无法从文档获取页面设置，使用A4默认值")
        page_settings = {
            'w': PAGE_W_PORTRAIT, 'h': PAGE_H_PORTRAIT,
            'margin_top': '1587', 'margin_bottom': '1587',
            'margin_left': '1417', 'margin_right': '1417',
            'margin_header': '1134', 'margin_footer': '1134',
            'margin_gutter': '0',
        }

    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        print("  警告：未找到文档body元素")
        return 0

    # 计算横向页面可用宽度（用于缩放表格列宽）
    landscape_page_w = int(page_settings['h'])  # 横向时w=h
    margin_left = int(page_settings['margin_left'])
    margin_right = int(page_settings['margin_right'])
    landscape_avail_w = landscape_page_w - margin_left - margin_right

    # 收集所有需要插入分节符的位置
    # 格式: [(caption_idx, after_idx, tbl_elem, caption_keyword)]
    insert_tasks = []
    body_children = list(body)

    for caption_kw in landscape_captions:
        # 在body直接子元素中查找包含caption关键词的标题段落
        caption_idx = None
        for idx, elem in enumerate(body_children):
            if elem.tag != f'{{{w_ns}}}p':
                continue
            text = _get_paragraph_text(elem, w_ns)
            if caption_kw in text:
                # 额外验证：标题段落通常以"表X.X"开头
                if re.match(r'^表\d+\.\d+\s', text):
                    caption_idx = idx
                    break

        if caption_idx is None:
            print(f"  警告：未找到含「{caption_kw}」的标题段落")
            continue

        # 从标题段落往下查找最近的表格
        tbl_elem = None
        tbl_idx = None
        for k in range(caption_idx + 1, len(body_children)):
            if body_children[k].tag == f'{{{w_ns}}}tbl':
                tbl_elem = body_children[k]
                tbl_idx = k
                break
            if body_children[k].tag == f'{{{w_ns}}}p':
                continue  # 可能是英文标题段落
            break

        if tbl_elem is None:
            print(f"  警告：「{caption_kw}」标题后未找到表格元素")
            continue

        # 确定landscape sectPr的插入位置（表格之后，考虑表注）
        after_idx = tbl_idx
        for k in range(tbl_idx + 1, min(tbl_idx + 3, len(body_children))):
            elem = body_children[k]
            if elem.tag != f'{{{w_ns}}}p':
                break
            text = _get_paragraph_text(elem, w_ns)
            if text.startswith('注：') or text.startswith('注:'):
                after_idx = k
            break

        # 向后查找表格后面最近的已有sectPr（定义当前节属性，含正确的headerRef）
        # 必须在插入新sectPr之前查找，否则会找到自己插入的
        ref_sect_pr = None
        for k in range(after_idx + 1, len(body_children)):
            elem = body_children[k]
            if elem.tag == f'{{{w_ns}}}p':
                sect = elem.find(f'.//{{{w_ns}}}sectPr')
                if sect is not None:
                    ref_sect_pr = sect
                    break
            elif elem.tag == f'{{{w_ns}}}sectPr':
                ref_sect_pr = elem
                break

        if ref_sect_pr is None:
            # fallback: body末尾的sectPr
            for child in reversed(body_children):
                if child.tag == f'{{{w_ns}}}sectPr':
                    ref_sect_pr = child
                    break

        insert_tasks.append((caption_idx, after_idx, tbl_elem, caption_kw, ref_sect_pr))

    # 按caption_idx倒序排列，从后往前插入避免索引偏移
    insert_tasks.sort(key=lambda x: x[0], reverse=True)

    processed = 0
    for caption_idx, after_idx, tbl_elem, caption_kw, ref_sect_pr in insert_tasks:
        # 1. 缩放表格列宽以匹配横向页面
        _scale_table_to_page_width(tbl_elem, landscape_avail_w, w_ns)

        # 2. 先插入后面的（landscape sectPr），再插入前面的（portrait sectPr）
        landscape_break = _create_section_break_para(
            w_ns, ref_sect_pr, page_settings, landscape=True)
        body.insert(after_idx + 1, landscape_break)

        portrait_break = _create_section_break_para(
            w_ns, ref_sect_pr, page_settings, landscape=False)
        body.insert(caption_idx, portrait_break)

        processed += 1
        print(f"  「{caption_kw}」→ 横向页面")

    return processed
