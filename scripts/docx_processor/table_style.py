"""
表格样式处理模块
处理表格标题：
1. 将标题从表格上方移到表格下方
2. 格式化为中英文双标题
3. 应用正确的样式和间距
"""
import xml.etree.ElementTree as ET
import re
import os
from .utils import NAMESPACES, append_text_with_math
from .config import STYLES, DEFAULT_QMD_FILENAME, EXCLUDE_CHAPTER_TITLES, THREE_LINE_BORDERS


def _ensure_child(parent, tag, ns, index=0):
    """确保子元素存在，不存在则创建
    
    Args:
        parent: 父元素
        tag: 子元素标签名（不含命名空间前缀）
        ns: 命名空间URI
        index: 插入位置（默认0，即最前面）
    
    Returns:
        子元素
    """
    full_tag = f'{{{ns}}}{tag}'
    child = parent.find(full_tag)
    if child is None:
        child = ET.Element(full_tag)
        parent.insert(index, child)
    return child


def load_en_table_captions_from_qmd(qmd_path=None):
    """从thesis.qmd读取英文表格标题
    
    支持两种表格格式：
    1. 简单Markdown表格: ": 标题 {#tbl-xxx}" 后紧跟 "<!-- tbl-cap-en: ... -->"
    2. 复杂HTML表格: "::: {#tbl-xxx}" 开头，":::" 结束后紧跟 "<!-- tbl-cap-en: ... -->"
    """
    if qmd_path is None:
        qmd_path = DEFAULT_QMD_FILENAME
    
    captions = {}
    if not os.path.exists(qmd_path):
        return captions
    
    with open(qmd_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    for i, line in enumerate(lines):
        match = re.match(r'^\s*<!-- tbl-cap-en:\s*(.+?)\s*-->', line)
        if not match:
            continue
        en_caption = match.group(1)
        # 向上查找最近的 {#tbl-xxx}（在 ::: {#tbl-xxx} 或 : 标题 {#tbl-xxx} 中）
        for j in range(i - 1, max(i - 500, -1), -1):
            tbl_match = re.search(r'\{#(tbl-[^}]+)\}', lines[j])
            if tbl_match:
                captions[tbl_match.group(1)] = en_caption
                break
    
    return captions


def load_table_notes_from_qmd(qmd_path=None):
    """从thesis.qmd读取表注
    
    格式: <!-- tbl-note: 注释内容 --> 放在表格标题后面
    """
    if qmd_path is None:
        qmd_path = DEFAULT_QMD_FILENAME
    
    notes = {}
    if not os.path.exists(qmd_path):
        return notes

    with open(qmd_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for i, line in enumerate(lines):
        match = re.match(r'^\s*<!-- tbl-note:\s*(.+?)\s*-->', line)
        if not match:
            continue

        note = match.group(1).strip()

        # 向上查找最近的 {#tbl-xxx}，兼容 pipe table 与 HTML table 块
        for j in range(i - 1, max(i - 500, -1), -1):
            tbl_match = re.search(r'\{#(tbl-[^}]+)\}', lines[j])
            if tbl_match:
                notes[tbl_match.group(1)] = note
                break

    return notes


def _extract_alignment_from_attrs(attrs_text):
    """从HTML标签属性中提取 text-align/align（仅支持 left/center）"""
    # 优先解析 style="text-align: ..."
    style_match = re.search(r'style\s*=\s*["\']([^"\']+)["\']', attrs_text, re.IGNORECASE)
    if style_match:
        style_text = style_match.group(1)
        align_match = re.search(r'text-align\s*:\s*(left|center)\b', style_text, re.IGNORECASE)
        if align_match:
            return align_match.group(1).lower()

    # 兼容 align="left|center"
    align_match = re.search(r'\balign\s*=\s*["\']?(left|center)["\']?', attrs_text, re.IGNORECASE)
    if align_match:
        return align_match.group(1).lower()

    return None


def _parse_html_table_column_alignments(html_text):
    """从HTML table代码中解析列对齐配置（0-based列索引）"""
    col_alignment = {}

    # 仅处理首个 table 块
    table_match = re.search(r'<table\b[^>]*>(.*?)</table>', html_text, re.IGNORECASE | re.DOTALL)
    if not table_match:
        return col_alignment
    table_html = table_match.group(1)

    # 1) 优先读取 <col> 的对齐设置
    col_index = 0
    for col_match in re.finditer(r'<col\b([^>]*)>', table_html, re.IGNORECASE):
        attrs_text = col_match.group(1)
        align = _extract_alignment_from_attrs(attrs_text)
        if align in ('left', 'center'):
            col_alignment[col_index] = align
        col_index += 1

    # 2) 再读取 th/td 的对齐设置，填充未指定列
    row_matches = re.finditer(r'<tr\b[^>]*>(.*?)</tr>', table_html, re.IGNORECASE | re.DOTALL)
    for row_match in row_matches:
        row_html = row_match.group(1)
        cursor = 0
        for cell_match in re.finditer(r'<(th|td)\b([^>]*)>(.*?)</\1>', row_html, re.IGNORECASE | re.DOTALL):
            attrs_text = cell_match.group(2)
            align = _extract_alignment_from_attrs(attrs_text)

            colspan = 1
            colspan_match = re.search(r'\bcolspan\s*=\s*["\']?(\d+)["\']?', attrs_text, re.IGNORECASE)
            if colspan_match:
                colspan = max(1, int(colspan_match.group(1)))

            if align in ('left', 'center'):
                for k in range(cursor, cursor + colspan):
                    if k not in col_alignment:
                        col_alignment[k] = align

            cursor += colspan

    return col_alignment


def load_table_column_alignments_from_qmd(qmd_path=None):
    """从QMD中的HTML表格代码读取列对齐设置

    解析结构：
    ::: {#tbl-xxx}
    ```{=html}
    <table>...</table>
    ```
    :::
    """
    if qmd_path is None:
        qmd_path = DEFAULT_QMD_FILENAME

    alignments = {}
    if not os.path.exists(qmd_path):
        return alignments

    with open(qmd_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    i = 0
    current_tbl_id = None
    while i < len(lines):
        line = lines[i]

        tbl_match = re.match(r'^\s*:::\s*\{#(tbl-[^}\s]+).*\}\s*$', line)
        if tbl_match:
            current_tbl_id = tbl_match.group(1)
            i += 1
            continue

        if current_tbl_id and re.match(r'^\s*```\{=html\}\s*$', line):
            i += 1
            html_lines = []
            while i < len(lines) and not re.match(r'^\s*```\s*$', lines[i]):
                html_lines.append(lines[i])
                i += 1

            html_text = ''.join(html_lines)
            col_alignment = _parse_html_table_column_alignments(html_text)
            if col_alignment:
                alignments[current_tbl_id] = col_alignment

            # 跳过结束 ``` 行
            if i < len(lines):
                i += 1
            continue

        if current_tbl_id and re.match(r'^\s*:::\s*$', line):
            current_tbl_id = None

        i += 1

    return alignments


# ---------------------------------------------------------------------------
# 自定义边框支持：在三线表基础上，允许用户通过 HTML CSS border 声明额外边框
# ---------------------------------------------------------------------------

def _parse_css_border_value(css_value):
    """解析单个 CSS border 值，如 '1px solid black'

    Returns:
        dict {'val', 'sz', 'color', 'space'} 或 None（none/无效时）
    """
    css_value = css_value.strip().lower()
    if not css_value or css_value == 'none' or css_value == '0':
        return None

    # 宽度
    width_pt = 0.5
    w_match = re.search(r'([\d.]+)\s*(px|pt)', css_value)
    if w_match:
        value = float(w_match.group(1))
        unit = w_match.group(2)
        width_pt = value * 0.75 if unit == 'px' else value
    elif 'thin' in css_value:
        width_pt = 0.5
    elif 'thick' in css_value:
        width_pt = 1.5
    elif 'medium' in css_value:
        width_pt = 1.0

    # 线型
    val = 'single'
    if 'dashed' in css_value:
        val = 'dashed'
    elif 'dotted' in css_value:
        val = 'dotted'
    elif 'double' in css_value:
        val = 'double'

    # 颜色
    color = '000000'
    c_match = re.search(r'#([0-9a-fA-F]{6})', css_value)
    if c_match:
        color = c_match.group(1).upper()
    else:
        c_match = re.search(r'#([0-9a-fA-F]{3})\b', css_value)
        if c_match:
            c = c_match.group(1)
            color = (c[0]*2 + c[1]*2 + c[2]*2).upper()

    sz = str(max(1, round(width_pt * 8)))
    return {'val': val, 'sz': sz, 'color': color, 'space': '0'}


def _parse_border_styles_from_attrs(attrs_text):
    """从 HTML 标签属性的 style 中提取 border 声明

    Returns:
        dict: 只包含声明了的方向，如 {'top': border_info, 'left': border_info}
    """
    borders = {}
    style_match = re.search(r'style\s*=\s*["\']([^"\']+)["\']', attrs_text, re.IGNORECASE)
    if not style_match:
        return borders

    style_text = style_match.group(1)

    # 简写 border（四个方向）
    all_match = re.search(r'(?<![a-z-])border\s*:\s*([^;]+)', style_text, re.IGNORECASE)
    if all_match:
        info = _parse_css_border_value(all_match.group(1))
        if info:
            for d in ('top', 'bottom', 'left', 'right'):
                borders[d] = dict(info)

    # 各方向独立声明（覆盖简写）
    for direction in ('top', 'bottom', 'left', 'right'):
        m = re.search(rf'border-{direction}\s*:\s*([^;]+)', style_text, re.IGNORECASE)
        if m:
            info = _parse_css_border_value(m.group(1))
            if info:
                borders[direction] = info
            else:
                borders.pop(direction, None)

    return borders


def _parse_html_table_cell_borders(html_text):
    """解析 HTML 表格中所有单元格的自定义边框声明

    支持三个层级的声明（优先级递增）：
    1. <col> 上的 border（列级别，仅 left/right 有效）
    2. <tr> 上的 border（行级别）
    3. <td>/<th> 上的 border（单元格级别）

    Returns:
        (cell_borders, header_row_count)
        cell_borders: {(row_idx, col_idx): {'top': info, ...}}
        header_row_count: <thead> 中的行数
    """
    cell_borders = {}

    table_match = re.search(r'<table\b[^>]*>(.*?)</table>', html_text,
                            re.IGNORECASE | re.DOTALL)
    if not table_match:
        return cell_borders, 0

    table_html = table_match.group(1)

    # thead 行数
    header_row_count = 0
    thead_match = re.search(r'<thead\b[^>]*>(.*?)</thead>', table_html,
                            re.IGNORECASE | re.DOTALL)
    if thead_match:
        header_row_count = len(re.findall(r'<tr\b', thead_match.group(1), re.IGNORECASE))

    # ---- 1. 解析 <col> 边框 ----
    col_borders = {}  # {col_idx: {'left': info, 'right': info}}
    ci = 0
    for col_m in re.finditer(r'<col\b([^>]*)/?>', table_html, re.IGNORECASE):
        attrs = col_m.group(1)
        span = 1
        span_m = re.search(r'\bspan\s*=\s*["\']?(\d+)', attrs, re.IGNORECASE)
        if span_m:
            span = max(1, int(span_m.group(1)))
        b = _parse_border_styles_from_attrs(attrs)
        for k in range(ci, ci + span):
            entry = {}
            if 'left' in b and k == ci:
                entry['left'] = b['left']
            if 'right' in b and k == ci + span - 1:
                entry['right'] = b['right']
            if entry:
                col_borders[k] = entry
        ci += span

    # ---- 2. 遍历行和单元格 ----
    grid = {}  # 占位网格
    row_level_borders = {}  # {row_idx: row_b} 记录行级 border，用于后处理
    row_idx = 0
    for row_m in re.finditer(r'<tr\b([^>]*)>(.*?)</tr>', table_html,
                             re.IGNORECASE | re.DOTALL):
        row_attrs = row_m.group(1)
        row_html = row_m.group(2)
        row_b = _parse_border_styles_from_attrs(row_attrs)

        if row_b and ('top' in row_b or 'bottom' in row_b):
            row_level_borders[row_idx] = row_b

        col_idx = 0
        for cell_m in re.finditer(r'<(th|td)\b([^>]*)>', row_html, re.IGNORECASE):
            while (row_idx, col_idx) in grid:
                col_idx += 1

            attrs = cell_m.group(2)
            colspan = 1
            rowspan = 1
            cs = re.search(r'\bcolspan\s*=\s*["\']?(\d+)', attrs, re.IGNORECASE)
            rs = re.search(r'\browspan\s*=\s*["\']?(\d+)', attrs, re.IGNORECASE)
            if cs:
                colspan = max(1, int(cs.group(1)))
            if rs:
                rowspan = max(1, int(rs.group(1)))

            cell_b = _parse_border_styles_from_attrs(attrs)

            # 标记占位 & 分配边框
            for r in range(row_idx, row_idx + rowspan):
                for c in range(col_idx, col_idx + colspan):
                    grid[(r, c)] = True
                    borders = {}

                    # 列级别默认
                    if c in col_borders:
                        borders.update(col_borders[c])

                    # 行级别 left/right（仅在单元格实际声明所在行应用）
                    if r == row_idx:
                        for d in ('left', 'right'):
                            if d in row_b:
                                borders[d] = row_b[d]

                    # 单元格级别
                    if 'top' in cell_b and r == row_idx:
                        borders['top'] = cell_b['top']
                    if 'bottom' in cell_b and r == row_idx + rowspan - 1:
                        borders['bottom'] = cell_b['bottom']
                    if 'left' in cell_b and c == col_idx:
                        borders['left'] = cell_b['left']
                    if 'right' in cell_b and c == col_idx + colspan - 1:
                        borders['right'] = cell_b['right']

                    if borders:
                        cell_borders.setdefault((r, c), {}).update(borders)

            col_idx += colspan
        row_idx += 1

    # ---- 3. 后处理：将 <tr> 的 border-top/bottom 应用到该行的所有逻辑列 ----
    # 这确保被 rowspan 占据的位置也能获得行级横线
    # 使用 setdefault 保证单元格级别声明的优先级更高
    if row_level_borders and grid:
        max_col = max(c for (_, c) in grid)
        for ri, rb in row_level_borders.items():
            for c in range(max_col + 1):
                if (ri, c) not in grid:
                    continue
                if 'bottom' in rb:
                    cell_borders.setdefault((ri, c), {}).setdefault('bottom', rb['bottom'])
                if 'top' in rb:
                    cell_borders.setdefault((ri, c), {}).setdefault('top', rb['top'])

    return cell_borders, header_row_count


def load_table_cell_borders_from_qmd(qmd_path=None):
    """从 QMD 的 HTML 表格中加载自定义边框配置

    Returns:
        dict: {tbl_id: (cell_borders, header_row_count)}
    """
    if qmd_path is None:
        qmd_path = DEFAULT_QMD_FILENAME

    result = {}
    if not os.path.exists(qmd_path):
        return result

    with open(qmd_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    i = 0
    current_tbl_id = None
    while i < len(lines):
        line = lines[i]

        tbl_match = re.match(r'^\s*:::\s*\{#(tbl-[^}\s]+).*\}\s*$', line)
        if tbl_match:
            current_tbl_id = tbl_match.group(1)
            i += 1
            continue

        if current_tbl_id and re.match(r'^\s*```\{=html\}\s*$', line):
            i += 1
            html_lines = []
            while i < len(lines) and not re.match(r'^\s*```\s*$', lines[i]):
                html_lines.append(lines[i])
                i += 1

            html_text = ''.join(html_lines)
            cell_borders, header_rows = _parse_html_table_cell_borders(html_text)
            if cell_borders:
                result[current_tbl_id] = (cell_borders, header_rows)

            if i < len(lines):
                i += 1
            continue

        if current_tbl_id and re.match(r'^\s*:::\s*$', line):
            current_tbl_id = None

        i += 1

    return result


def apply_table_cell_borders(tbl, cell_borders, header_row_count, w_ns):
    """将自定义边框应用到 DOCX 表格（在三线表基础上叠加）

    对有自定义边框的表格，为每个单元格写入完整的 tcBorders：
    - 三线表基础边框（顶线、表头底线、底线）
    - 用户声明的额外边框
    - 其他方向设为 nil

    Args:
        tbl: w:tbl 元素
        cell_borders: {(row_idx, col_idx): {'top': info, ...}}
        header_row_count: 表头行数
        w_ns: Word 命名空间

    Returns:
        应用了边框的单元格数
    """
    rows = list(tbl.findall(f'{{{w_ns}}}tr'))
    total_rows = len(rows)
    if total_rows == 0:
        return 0

    applied = 0
    for row_idx, tr in enumerate(rows):
        col_idx = 0
        for tc in tr.findall(f'{{{w_ns}}}tc'):
            # 计算 colspan
            colspan = 1
            tcPr = tc.find(f'{{{w_ns}}}tcPr')
            if tcPr is not None:
                gs = tcPr.find(f'{{{w_ns}}}gridSpan')
                if gs is not None:
                    sv = gs.get(f'{{{w_ns}}}val')
                    if sv and sv.isdigit():
                        colspan = max(1, int(sv))

            # ---- 构建该单元格的完整边框 ----
            borders = {'top': None, 'bottom': None, 'left': None, 'right': None}

            # 三线表基础边框
            if row_idx == 0:
                borders['top'] = dict(THREE_LINE_BORDERS['top'])
            if header_row_count > 0 and row_idx == header_row_count - 1:
                borders['bottom'] = dict(THREE_LINE_BORDERS['header_bottom'])
            if row_idx == total_rows - 1:
                borders['bottom'] = dict(THREE_LINE_BORDERS['bottom'])

            # 叠加自定义边框（检查该物理单元格覆盖的所有逻辑列）
            for c in range(col_idx, col_idx + colspan):
                if (row_idx, c) not in cell_borders:
                    continue
                custom = cell_borders[(row_idx, c)]
                if 'top' in custom:
                    borders['top'] = custom['top']
                if 'bottom' in custom:
                    borders['bottom'] = custom['bottom']
                if 'left' in custom and c == col_idx:
                    borders['left'] = custom['left']
                if 'right' in custom and c == col_idx + colspan - 1:
                    borders['right'] = custom['right']

            # ---- 写入 tcBorders ----
            if tcPr is None:
                tcPr = ET.Element(f'{{{w_ns}}}tcPr')
                tc.insert(0, tcPr)

            old_b = tcPr.find(f'{{{w_ns}}}tcBorders')
            if old_b is not None:
                tcPr.remove(old_b)

            tcBorders = ET.SubElement(tcPr, f'{{{w_ns}}}tcBorders')
            for direction in ('top', 'left', 'bottom', 'right'):
                elem = ET.SubElement(tcBorders, f'{{{w_ns}}}{direction}')
                if borders[direction] is not None:
                    for attr, value in borders[direction].items():
                        elem.set(f'{{{w_ns}}}{attr}', value)
                else:
                    elem.set(f'{{{w_ns}}}val', 'nil')

            applied += 1
            col_idx += colspan

    return applied


def create_caption_paragraph(w_ns, text, style_id, center=True):
    """
    创建一个干净的题注段落
    显式设置段前段后间距为0
    支持 $...$ 中的LaTeX公式，转换为OMML数学元素
    
    Args:
        w_ns: Word命名空间
        text: 段落文本（可含$...$公式）
        style_id: 样式ID
        center: 是否居中（默认True）
    """
    para = ET.Element(f'{{{w_ns}}}p')
    
    # 段落属性
    pPr = ET.SubElement(para, f'{{{w_ns}}}pPr')
    pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
    pStyle.set(f'{{{w_ns}}}val', style_id)
    
    # 显式设置段前段后间距为0
    spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
    spacing.set(f'{{{w_ns}}}after', '0')
    spacing.set(f'{{{w_ns}}}before', '0')
    
    if center:
        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
        jc.set(f'{{{w_ns}}}val', 'center')
    
    # 文本内容（支持 $...$ 公式）
    append_text_with_math(para, w_ns, text)
    
    return para


LATEX_TO_UNICODE = {
    r'\alpha': '\u03b1', r'\beta': '\u03b2', r'\gamma': '\u03b3', r'\delta': '\u03b4',
    r'\epsilon': '\u03b5', r'\zeta': '\u03b6', r'\eta': '\u03b7', r'\theta': '\u03b8',
    r'\iota': '\u03b9', r'\kappa': '\u03ba', r'\lambda': '\u03bb', r'\mu': '\u03bc',
    r'\nu': '\u03bd', r'\xi': '\u03be', r'\pi': '\u03c0', r'\rho': '\u03c1',
    r'\sigma': '\u03c3', r'\tau': '\u03c4', r'\upsilon': '\u03c5', r'\phi': '\u03c6',
    r'\chi': '\u03c7', r'\psi': '\u03c8', r'\omega': '\u03c9',
    r'\Alpha': '\u0391', r'\Beta': '\u0392', r'\Gamma': '\u0393', r'\Delta': '\u0394',
    r'\Theta': '\u0398', r'\Lambda': '\u039b', r'\Pi': '\u03a0', r'\Sigma': '\u03a3',
    r'\Phi': '\u03a6', r'\Psi': '\u03a8', r'\Omega': '\u03a9',
    r'\times': '\u00d7', r'\pm': '\u00b1', r'\leq': '\u2264', r'\geq': '\u2265',
    r'\neq': '\u2260', r'\approx': '\u2248', r'\infty': '\u221e',
}


def _latex_to_omml_text(latex_str):
    """将简单的LaTeX公式字符串转换为用于OMML的Unicode文本"""
    result = latex_str.strip()
    for cmd, char in sorted(LATEX_TO_UNICODE.items(), key=lambda x: -len(x[0])):
        result = result.replace(cmd, char)
    return result


def _create_text_run(w_ns, text, font_size='21'):
    """创建带字体属性的文本run"""
    run = ET.Element(f'{{{w_ns}}}r')
    rPr = ET.SubElement(run, f'{{{w_ns}}}rPr')
    rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
    rFonts.set(f'{{{w_ns}}}ascii', 'Times New Roman')
    rFonts.set(f'{{{w_ns}}}eastAsia', '宋体')
    rFonts.set(f'{{{w_ns}}}hAnsi', 'Times New Roman')
    sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
    sz.set(f'{{{w_ns}}}val', font_size)
    szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
    szCs.set(f'{{{w_ns}}}val', font_size)
    t = ET.SubElement(run, f'{{{w_ns}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return run


def _create_omath_element(m_ns, w_ns, latex_content):
    """将LaTeX内容转换为OMML oMath元素"""
    omath = ET.Element(f'{{{m_ns}}}oMath')
    math_run = ET.SubElement(omath, f'{{{m_ns}}}r')
    # 设置数学字体
    mr_pr = ET.SubElement(math_run, f'{{{w_ns}}}rPr')
    mr_fonts = ET.SubElement(mr_pr, f'{{{w_ns}}}rFonts')
    mr_fonts.set(f'{{{w_ns}}}ascii', 'Cambria Math')
    mr_fonts.set(f'{{{w_ns}}}hAnsi', 'Cambria Math')
    mt = ET.SubElement(math_run, f'{{{m_ns}}}t')
    mt.text = _latex_to_omml_text(latex_content)
    return omath


def create_table_note_paragraph(w_ns, text):
    """
    创建表注段落
    格式：宋体5号（14磅行间距，左对齐
    支持 $...$ 中的LaTeX公式，转换为OMML数学元素
    
    Args:
        w_ns: Word命名空间
        text: 表注文本（可含$...$公式）
    """
    m_ns = NAMESPACES['m']
    para = ET.Element(f'{{{w_ns}}}p')
    
    # 段落属性
    pPr = ET.SubElement(para, f'{{{w_ns}}}pPr')
    
    # 14磅行间距 = 280 twips (1磅 = 20 twips)
    spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
    spacing.set(f'{{{w_ns}}}line', '280')
    spacing.set(f'{{{w_ns}}}lineRule', 'exact')
    spacing.set(f'{{{w_ns}}}after', '0')
    spacing.set(f'{{{w_ns}}}before', '0')
    
    # 左对齐
    jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
    jc.set(f'{{{w_ns}}}val', 'left')
    
    # 按 $...$ 分割文本，交替生成普通文本run和OMML数学元素
    parts = re.split(r'(\$[^$]+\$)', text)
    
    # 去除公式前后相邻文本段的空格
    for i in range(len(parts)):
        if not parts[i]:
            continue
        if parts[i].startswith('$') and parts[i].endswith('$'):
            continue
        for j in range(i - 1, -1, -1):
            if parts[j]:
                if parts[j].startswith('$') and parts[j].endswith('$'):
                    parts[i] = parts[i].lstrip(' ')
                break
        for j in range(i + 1, len(parts)):
            if parts[j]:
                if parts[j].startswith('$') and parts[j].endswith('$'):
                    parts[i] = parts[i].rstrip(' ')
                break
    
    for part in parts:
        if not part:
            continue
        if part.startswith('$') and part.endswith('$'):
            # 公式部分：去掉首尾$，转为OMML
            latex_content = part[1:-1]
            omath = _create_omath_element(m_ns, w_ns, latex_content)
            para.append(omath)
        else:
            # 普通文本
            run = _create_text_run(w_ns, part)
            para.append(run)
    
    return para


def get_paragraph_text(para, w_ns):
    """获取段落的纯文本内容"""
    text = ''
    for t in para.iter(f'{{{w_ns}}}t'):
        if t.text:
            text += t.text
    return text


def find_parent(root, target):
    """查找元素的父节点"""
    for parent in root.iter():
        for child in parent:
            if child is target:
                return parent
    return None


def _extract_nested_table_caption(root, w_ns):
    """处理 HTML 表格的嵌套结构：将标题从表格内部移到外部
    
    Pandoc 将 HTML 表格的 caption 转换为嵌套结构：
    - 外层表格只有一个单元格
    - 单元格内包含：书签 + 标题段落 + 嵌套的数据表格
    
    此函数将这种结构转换为：
    - 标题段落（在表格外部，包含书签）
    - 数据表格
    """
    extracted_count = 0
    
    # 遍历所有表格
    for tbl in list(root.iter(f'{{{w_ns}}}tbl')):
        # 检查是否是包含嵌套表格的外层表格
        rows = list(tbl.findall(f'{{{w_ns}}}tr'))
        if len(rows) != 1:
            continue
        
        cells = list(rows[0].findall(f'{{{w_ns}}}tc'))
        if len(cells) != 1:
            continue
        
        cell = cells[0]
        
        # 检查单元格内是否有嵌套表格
        nested_tables = list(cell.findall(f'{{{w_ns}}}tbl'))
        if not nested_tables:
            continue
        
        # 收集所有元素，区分书签、标题段落和其他元素
        bookmarks = []  # 书签元素（bookmarkStart 和 bookmarkEnd）
        caption_paras = []
        nested_tbls = []
        other_elements = []
        
        for child in list(cell):
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if tag_local == 'bookmarkStart':
                name = child.get(f'{{{w_ns}}}name', '')
                if name.startswith('tbl-'):
                    bookmarks.append(child)
            elif tag_local == 'bookmarkEnd':
                # 检查是否对应 tbl- 书签
                bookmark_id = child.get(f'{{{w_ns}}}id', '')
                # 查找对应的 bookmarkStart
                for bm in bookmarks:
                    if bm.get(f'{{{w_ns}}}id', '') == bookmark_id:
                        bookmarks.append(child)
                        break
            elif child.tag == f'{{{w_ns}}}p':
                text = get_paragraph_text(child, w_ns)
                if re.match(r'^Table\s+\d+:', text):
                    caption_paras.append(child)
                else:
                    other_elements.append(child)
            elif child.tag == f'{{{w_ns}}}tbl':
                nested_tbls.append(child)
            elif tag_local == 'tcPr':
                pass  # 跳过单元格属性
            else:
                other_elements.append(child)
        
        if not caption_paras:
            continue
        
        # 找到外层表格的父节点
        parent = find_parent(root, tbl)
        if parent is None:
            continue
        
        # 获取外层表格的位置
        tbl_idx = list(parent).index(tbl)
        
        # 移除外层表格
        parent.remove(tbl)
        
        # 在原位置插入：书签 + 标题段落 + 嵌套表格
        insert_idx = tbl_idx
        
        # 将书签插入到第一个标题段落内部（段落开头）
        if bookmarks and caption_paras:
            first_para = caption_paras[0]
            # 在段落开头插入书签
            for i, bm in enumerate(bookmarks):
                first_para.insert(i, bm)
        
        # 插入标题段落
        for para in caption_paras:
            parent.insert(insert_idx, para)
            insert_idx += 1
        
        # 插入嵌套表格
        for tbl_elem in nested_tbls:
            parent.insert(insert_idx, tbl_elem)
            insert_idx += 1
        
        extracted_count += 1
    
    return extracted_count


def process_tables(root, qmd_path=None):
    """处理表格样式和题注"""
    print("正在处理表格样式...")
    
    # 从qmd文件读取英文标题、表注、列对齐和自定义边框
    en_captions = load_en_table_captions_from_qmd(qmd_path)
    table_notes = load_table_notes_from_qmd(qmd_path)
    table_col_alignments = load_table_column_alignments_from_qmd(qmd_path)
    table_cell_borders = load_table_cell_borders_from_qmd(qmd_path)
    print(
        f"  从qmd文件读取到 {len(en_captions)} 个英文表格标题，"
        f"{len(table_notes)} 个表注，{len(table_col_alignments)} 个列对齐配置，"
        f"{len(table_cell_borders)} 个自定义边框配置"
    )
    
    w_ns = NAMESPACES['w']
    
    # 首先处理 HTML 表格的嵌套结构，将标题移到表格外部
    extracted = _extract_nested_table_caption(root, w_ns)
    if extracted > 0:
        print(f"  从 {extracted} 个嵌套表格中提取了标题")
    
    # 查找所有tbl-书签，建立顺序映射（从root全局搜索）
    tbl_bookmarks = []
    for bm in root.iter(f'{{{w_ns}}}bookmarkStart'):
        name = bm.get(f'{{{w_ns}}}name', '')
        if name.startswith('tbl-'):
            tbl_bookmarks.append(name)
    
    # 动态检测章节并收集表格标题
    # 遍历文档段落，跟踪当前章节号，为每个表格分配正确的章节
    current_chapter = 0
    caption_paras_info = []  # (para, tbl_num_orig, cn_caption, direct_tbl_id, chapter)
    
    for para in root.iter(f'{{{w_ns}}}p'):
        # 检测章节标题（Heading1）
        pPr = para.find(f'{{{w_ns}}}pPr')
        if pPr is not None:
            pStyle = pPr.find(f'{{{w_ns}}}pStyle')
            if pStyle is not None:
                style_val = pStyle.get(f'{{{w_ns}}}val', '')
                if style_val in (STYLES['heading1'], STYLES.get('heading1_alt', 'Heading1')):
                    para_text = get_paragraph_text(para, w_ns).strip()
                    para_text_nospace = para_text.replace(' ', '').replace('\u3000', '')
                    is_excluded = any(exclude in para_text_nospace for exclude in EXCLUDE_CHAPTER_TITLES)
                    if not is_excluded:
                        current_chapter += 1
        
        text = get_paragraph_text(para, w_ns)
        
        # 匹配 Quarto 生成的表格标题格式
        # 格式1: "Table X: 标题"
        match = re.match(r'^Table\s*(\d+):\s*(.*)$', text)
        if match:
            tbl_num_orig = int(match.group(1))
            cn_caption = match.group(2)
            caption_paras_info.append((para, tbl_num_orig, cn_caption, None, current_chapter))
            continue
        
        # 格式2: ": 标题 {#tbl-xxx}" (Quarto pipe table 格式)
        match = re.match(r'^:\s*(.+?)\s*\{#(tbl-[^}]+)\}$', text)
        if match:
            cn_caption = match.group(1)
            tbl_id = match.group(2)
            caption_paras_info.append((para, None, cn_caption, tbl_id, current_chapter))
    
    print(f"  找到 {len(caption_paras_info)} 个表格")
    
    # 处理每个表格标题
    table_map = {}
    chapter_tbl_count = {}
    tbl_counter = 0
    pending_table_alignments = []  # (tbl_id, target_tbl)
    
    for para, tbl_num_orig, cn_caption, direct_tbl_id, chapter in caption_paras_info:
        tbl_counter += 1
        
        # 确定表格ID
        if direct_tbl_id:
            tbl_id = direct_tbl_id
        elif tbl_num_orig and tbl_num_orig <= len(tbl_bookmarks):
            tbl_id = tbl_bookmarks[tbl_num_orig - 1]
        else:
            continue
        
        chapter_tbl_count[chapter] = chapter_tbl_count.get(chapter, 0) + 1
        tbl_num = chapter_tbl_count[chapter]
        table_map[tbl_id] = (chapter, tbl_num)
        
        # 获取英文标题
        en_caption = en_captions.get(tbl_id, cn_caption)
        
        # 创建新的中文题注段落，使用"题注"样式
        cn_text = f"表{chapter}.{tbl_num} {cn_caption}"
        new_cn_para = create_caption_paragraph(w_ns, cn_text, STYLES['caption'])
        
        # 创建新的英文题注段落，使用"题注"样式
        en_text = f"Table {chapter}.{tbl_num} {en_caption}"
        new_en_para = create_caption_paragraph(w_ns, en_text, STYLES['caption'])
        
        # 找到原标题段落的父节点
        parent = find_parent(root, para)
        if parent is not None:
            # 获取原段落的位置
            idx = list(parent).index(para)
            
            # 移除原标题段落
            parent.remove(para)
            
            # 表格标题在表格上方，先插入英文再插入中文（这样中文在上）
            parent.insert(idx, new_en_para)
            parent.insert(idx, new_cn_para)

            # 找到当前标题对应的第一个表格
            target_tbl = None
            target_tbl_pos = None
            for j, elem in enumerate(list(parent)):
                if elem.tag == f'{{{w_ns}}}tbl' and j > idx:
                    target_tbl = elem
                    target_tbl_pos = j
                    break

            # 记录目标表格，后续在三线表样式应用后再统一应用列对齐
            if target_tbl is not None:
                pending_table_alignments.append((tbl_id, target_tbl))
            
            # 如果有表注，需要在表格下方插入
            # 找到表格元素并在其后插入表注
            if tbl_id in table_notes:
                note_text = f"注：{table_notes[tbl_id]}"
                note_para = create_table_note_paragraph(w_ns, note_text)
                if target_tbl_pos is not None:
                    # 在表格后面插入表注
                    parent.insert(target_tbl_pos + 1, note_para)
                    print(f"      添加表注: {table_notes[tbl_id][:20]}...")
        
        print(f"    表{chapter}.{tbl_num}: {cn_caption}")
    
    # 更新交叉引用
    update_table_references(root, table_map)
    
    # 为所有表格应用"三线表"样式
    apply_three_line_table_style(root)

    # 在三线表样式之后应用列对齐，避免段落样式处理覆盖对齐效果
    for tbl_id, target_tbl in pending_table_alignments:
        if tbl_id not in table_col_alignments:
            continue
        applied = apply_table_column_alignment(
            target_tbl,
            table_col_alignments[tbl_id],
            w_ns,
        )
        if applied > 0:
            print(f"      应用列对齐: {tbl_id}（{applied} 列）")

    # 在三线表样式之后应用自定义边框（在三线表基础上叠加用户声明的额外边框）
    for tbl_id, target_tbl in pending_table_alignments:
        if tbl_id not in table_cell_borders:
            continue
        cell_borders, header_rows = table_cell_borders[tbl_id]
        applied = apply_table_cell_borders(
            target_tbl,
            cell_borders,
            header_rows,
            w_ns,
        )
        if applied > 0:
            print(f"      应用自定义边框: {tbl_id}（{applied} 个单元格）")
    
    return table_map


def _table_contains_image(tbl, w_ns):
    """检查表格是否包含图片（用于布局的表格）"""
    # 检查是否包含 drawing 元素（图片）
    for drawing in tbl.iter(f'{{{w_ns}}}drawing'):
        return True
    # 检查是否包含 pict 元素（旧式图片）
    for pict in tbl.iter(f'{{{w_ns}}}pict'):
        return True
    return False


def apply_table_column_alignment(tbl, col_alignment, w_ns):
    """将列对齐配置应用到单个Word表格

    Args:
        tbl: w:tbl 元素
        col_alignment: dict[int, str]，值为 left/center
        w_ns: Word 命名空间

    Returns:
        实际应用了对齐的列数
    """
    if not col_alignment:
        return 0

    applied_cols = set()

    for tr in tbl.findall(f'{{{w_ns}}}tr'):
        col_idx = 0
        for tc in tr.findall(f'{{{w_ns}}}tc'):
            colspan = 1
            tcPr = tc.find(f'{{{w_ns}}}tcPr')
            if tcPr is not None:
                grid_span = tcPr.find(f'{{{w_ns}}}gridSpan')
                if grid_span is not None:
                    span_val = grid_span.get(f'{{{w_ns}}}val')
                    if span_val and span_val.isdigit():
                        colspan = max(1, int(span_val))

            align = None
            for k in range(col_idx, col_idx + colspan):
                if k in col_alignment:
                    align = col_alignment[k]
                    applied_cols.add(k)
                    break

            if align in ('left', 'center'):
                for para in tc.iter(f'{{{w_ns}}}p'):
                    pPr = para.find(f'{{{w_ns}}}pPr')
                    if pPr is None:
                        pPr = ET.Element(f'{{{w_ns}}}pPr')
                        para.insert(0, pPr)

                    # 先移除所有已有的对齐设置，避免出现多个 w:jc 冲突
                    for old_jc in list(pPr.findall(f'{{{w_ns}}}jc')):
                        pPr.remove(old_jc)

                    jc = ET.Element(f'{{{w_ns}}}jc')
                    pPr.append(jc)
                    jc.set(f'{{{w_ns}}}val', align)

            col_idx += colspan

    return len(applied_cols)


def _is_caption_paragraph(para, w_ns):
    """判断段落是否为表格标题（caption）段落
    
    标题段落的特征：
    1. 文本以"表X.Y"或"Table X.Y"开头
    2. 或者包含表格ID书签（tbl-xxx）
    """
    # 检查是否包含表格书签
    for bm in para.iter(f'{{{w_ns}}}bookmarkStart'):
        name = bm.get(f'{{{w_ns}}}name', '')
        if name.startswith('tbl-'):
            return True
    
    # 检查文本内容
    text = get_paragraph_text(para, w_ns)
    if re.match(r'^表\d+\.\d+\s', text) or re.match(r'^Table\s+\d+\.\d+\s', text):
        return True
    if re.match(r'^Table\s+\d+:\s', text):  # Quarto 原始格式
        return True
    
    return False


def apply_three_line_table_style(root):
    """为文档中的数据表格应用"三线表"样式和段落样式（跳过包含图片的布局表格）
    
    样式应用规则：
    - 表格标题段落：应用"题注"样式（caption）
    - 表格内容段落：应用"表"样式（table_content）
    - 表格本身：应用"三线表"设计样式
    """
    w_ns = NAMESPACES['w']
    table_style_id = STYLES['three_line_table']
    table_content_style_id = STYLES['table_content']  # "表"段落样式
    caption_style_id = STYLES['caption']  # "题注"段落样式
    
    tbl_count = 0
    skipped_count = 0
    for tbl in root.iter(f'{{{w_ns}}}tbl'):
        # 跳过包含图片的表格（这些是用于布局的表格）
        if _table_contains_image(tbl, w_ns):
            skipped_count += 1
            continue
        
        # 获取或创建 tblPr（表格属性）
        tblPr = _ensure_child(tbl, 'tblPr', w_ns, index=0)
        
        # 查找或创建 tblStyle 元素
        tblStyle = tblPr.find(f'{{{w_ns}}}tblStyle')
        if tblStyle is None:
            tblStyle = ET.Element(f'{{{w_ns}}}tblStyle')
            # tblStyle 应该在 tblPr 的最前面
            tblPr.insert(0, tblStyle)
        
        # 设置样式ID为"三线表"
        tblStyle.set(f'{{{w_ns}}}val', table_style_id)
        
        # 设置表格宽度为与窗口等宽（100%页面宽度）
        tblW = tblPr.find(f'{{{w_ns}}}tblW')
        if tblW is None:
            tblW = ET.SubElement(tblPr, f'{{{w_ns}}}tblW')
        tblW.set(f'{{{w_ns}}}w', '5000')
        tblW.set(f'{{{w_ns}}}type', 'pct')
        
        # 设置表格居中对齐
        jc = tblPr.find(f'{{{w_ns}}}jc')
        if jc is None:
            jc = ET.SubElement(tblPr, f'{{{w_ns}}}jc')
        jc.set(f'{{{w_ns}}}val', 'center')
        
        # 设置 tblLook 启用首行样式（三线表需要 firstRow="1" 才能正确显示表头下方的细线）
        tblLook = tblPr.find(f'{{{w_ns}}}tblLook')
        if tblLook is None:
            tblLook = ET.SubElement(tblPr, f'{{{w_ns}}}tblLook')
        tblLook.set(f'{{{w_ns}}}firstRow', '1')
        tblLook.set(f'{{{w_ns}}}lastRow', '0')
        tblLook.set(f'{{{w_ns}}}firstColumn', '0')
        tblLook.set(f'{{{w_ns}}}lastColumn', '0')
        tblLook.set(f'{{{w_ns}}}noHBand', '1')
        tblLook.set(f'{{{w_ns}}}noVBand', '1')
        
        # 清除表格级别的直接边框设置（让样式生效）
        tblBorders = tblPr.find(f'{{{w_ns}}}tblBorders')
        if tblBorders is not None:
            tblPr.remove(tblBorders)
        
        # 清除单元格的直接边框设置，并为表格段落应用适当的样式
        for tc in tbl.iter(f'{{{w_ns}}}tc'):
            tcPr = tc.find(f'{{{w_ns}}}tcPr')
            if tcPr is not None:
                tcBorders = tcPr.find(f'{{{w_ns}}}tcBorders')
                if tcBorders is not None:
                    tcPr.remove(tcBorders)
            
            # 为单元格内的每个段落应用适当的样式
            for para in tc.iter(f'{{{w_ns}}}p'):
                pPr = para.find(f'{{{w_ns}}}pPr')
                if pPr is None:
                    pPr = ET.Element(f'{{{w_ns}}}pPr')
                    para.insert(0, pPr)
                
                # 判断是否为标题段落
                is_caption = _is_caption_paragraph(para, w_ns)
                
                # 设置段落样式
                pStyle = pPr.find(f'{{{w_ns}}}pStyle')
                if pStyle is None:
                    pStyle = ET.Element(f'{{{w_ns}}}pStyle')
                    pPr.insert(0, pStyle)
                
                if is_caption:
                    # 标题段落使用"题注"样式
                    pStyle.set(f'{{{w_ns}}}val', caption_style_id)
                else:
                    # 内容段落使用"表"样式
                    pStyle.set(f'{{{w_ns}}}val', table_content_style_id)
        
        tbl_count += 1
    
    print(f"  为 {tbl_count} 个数据表格应用了\"三线表\"样式（跳过 {skipped_count} 个图片布局表格）")


def update_table_references(root, table_map):
    """更新表格交叉引用"""
    w_ns = NAMESPACES['w']
    refs_count = 0
    
    for para in root.iter(f'{{{w_ns}}}p'):
        children = list(para)
        to_remove = []
        
        for i, child in enumerate(children):
            if child.tag == f'{{{w_ns}}}hyperlink':
                anchor = child.get(f'{{{w_ns}}}anchor')
                if anchor and anchor in table_map:
                    ch, tn = table_map[anchor]
                    new_text = f"表{ch}.{tn}"
                    
                    runs = [r for r in child if r.tag == f'{{{w_ns}}}r']
                    if runs:
                        # 只保留第一个run
                        for r in runs[1:]:
                            child.remove(r)
                        
                        # 更新文本
                        t = runs[0].find(f'{{{w_ns}}}t')
                        if t is None:
                            t = ET.SubElement(runs[0], f'{{{w_ns}}}t')
                        t.text = new_text
                        refs_count += 1
                    
                    # 移除前面的"表"和空格
                    j = i - 1
                    while j >= 0:
                        prev = children[j]
                        if prev.tag == f'{{{w_ns}}}r':
                            t = prev.find(f'{{{w_ns}}}t')
                            if t is not None and t.text is not None:
                                text = t.text
                                if text.strip() == '':
                                    to_remove.append(prev)
                                    j -= 1
                                    continue
                                elif text.endswith('表'):
                                    t.text = text[:-1]
                                    break
                                elif text.endswith('表 '):
                                    t.text = text[:-2]
                                    break
                                elif text.strip() == '表':
                                    to_remove.append(prev)
                                    break
                                else:
                                    break
                            else:
                                break
                        else:
                            break
                        j -= 1
        
        for item in to_remove:
            try:
                para.remove(item)
            except:
                pass
    
    print(f"  更新了 {refs_count} 个表格交叉引用")
