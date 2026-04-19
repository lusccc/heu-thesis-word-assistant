"""
符号样式处理模块
处理QCA表格中的特殊符号，通过字号大小区分核心/边缘条件

在QCA（定性比较分析）论文中：
- ⬤ 核心条件存在（正常字号）
- ⊗ 核心条件缺失（正常字号）
- ● 边缘条件存在（缩小字号）— CJK字体中●为全角，需缩小以区分⬤
- ⊘ 边缘条件缺失 → 替换为⊗并缩小字号

核心/边缘条件通过符号大小区分。
"""
import copy
import xml.etree.ElementTree as ET
from .utils import NAMESPACES

# ⊘ (U+2298, CIRCLED DIVISION SLASH) → 缩小的 ⊗ (U+2297, CIRCLED TIMES)
SYMBOL_OSLASH = '\u2298'   # ⊘ 边缘条件缺失（源文件中的占位符）
SYMBOL_OTIMES = '\u2297'   # ⊗ 替换目标

# ● (U+25CF, BLACK CIRCLE) — 边缘条件存在
# 在CJK字体（宋体/等线）中为全角大圆，与⬤视觉大小几乎无区别，需缩小字号
SYMBOL_BULLET = '\u25CF'   # ● 边缘条件存在

# ⬤ (U+2B24, BLACK LARGE CIRCLE) — 核心条件存在
SYMBOL_LARGE_CIRCLE = '\u2B24'  # ⬤ 核心条件存在

# 缩小后的字号（半磅单位）
# 表格正文通常为五号（10.5pt = 21半磅），缩小版约 8pt = 16半磅
SHRUNK_FONT_SIZE_HALF_PT = 16
DEFAULT_FONT_SIZE_HALF_PT = 21  # 表格正文默认字号

# ⊗ 字形在字体中偏小，按百分比放大以匹配同级符号（⬤/●）
OTIMES_SCALE_PERCENT = 115  # 放大115%
OTIMES_CORE_HALF_PT = str(round(DEFAULT_FONT_SIZE_HALF_PT * OTIMES_SCALE_PERCENT / 100))  # 核心⊗
OTIMES_EDGE_HALF_PT = str(round(SHRUNK_FONT_SIZE_HALF_PT * OTIMES_SCALE_PERCENT / 100))   # 边缘⊗
SHRUNK_FONT_SIZE_HALF_PT = str(SHRUNK_FONT_SIZE_HALF_PT)  # 转为字符串

# 统一字体：解决不同QCA符号因Unicode范围差异被Word映射到不同字体
# （⬤→eastAsia全角 vs ⊗→西文半角）导致大小不一致的问题
QCA_SYMBOL_FONT = 'Segoe UI Symbol'


def process_circled_symbols(root):
    """处理QCA表格中的所有特殊符号

    边缘条件（替换/缩小字号 + 统一字体）：
    1. ⊘ → 缩小的⊗（替换字符 + 缩小字号 + 统一字体）
    2. ● → 缩小的●（缩小字号 + 统一字体）

    核心条件（仅统一字体，解决⬤全角 vs ⊗半角大小不一致）：
    3. ⬤ → 统一字体
    4. ⊗ → 统一字体（覆盖核心版，边缘版已有字号不受影响）

    处理方式：
    - run中仅含目标符号：直接操作
    - run中目标符号与其他文字混合：拆分run，仅对符号部分设置样式

    Args:
        root: 文档根元素

    Returns:
        int: 处理的符号总数
    """
    w_ns = NAMESPACES['w']
    font = QCA_SYMBOL_FONT
    shrunk = SHRUNK_FONT_SIZE_HALF_PT

    print(f"正在处理QCA符号（字体={font}, ⊗放大{OTIMES_SCALE_PERCENT}%）...")

    # --- 第一轮：⊘ → 缩小的⊗（替换 + 缩小 + 放大⊗ + 字体） ---
    parent_map = {id(child): parent for parent in root.iter() for child in parent}
    oslash_count = _process_symbol_replace(
        root, w_ns, parent_map,
        src_char=SYMBOL_OSLASH, dst_char=SYMBOL_OTIMES,
        font_size=OTIMES_EDGE_HALF_PT, font_name=font
    )
    print(f"  ⊘ → 边缘⊗: {oslash_count} 个（{int(OTIMES_EDGE_HALF_PT)//2}pt, {font}）")

    # --- 第二轮：● → 缩小的●（缩小 + 字体） ---
    parent_map = {id(child): parent for parent in root.iter() for child in parent}
    bullet_count = _process_symbol_replace(
        root, w_ns, parent_map,
        src_char=SYMBOL_BULLET, dst_char=SYMBOL_BULLET,
        font_size=shrunk, font_name=font
    )
    print(f"  ● → 边缘●: {bullet_count} 个（{int(shrunk)//2}pt, {font}）")

    # --- 第三轮：⬤ 统一字体（不改字号） ---
    parent_map = {id(child): parent for parent in root.iter() for child in parent}
    large_count = _process_symbol_replace(
        root, w_ns, parent_map,
        src_char=SYMBOL_LARGE_CIRCLE, dst_char=SYMBOL_LARGE_CIRCLE,
        font_name=font
    )
    print(f"  ⬤ 统一字体: {large_count} 个（{font}）")

    # --- 第四轮：⊗ 放大+字体（核心版设放大字号，边缘版已有字号则跳过） ---
    parent_map = {id(child): parent for parent in root.iter() for child in parent}
    otimes_count = _process_symbol_replace(
        root, w_ns, parent_map,
        src_char=SYMBOL_OTIMES, dst_char=SYMBOL_OTIMES,
        font_size=OTIMES_CORE_HALF_PT, font_name=font,
        skip_if_has_size=True
    )
    print(f"  ⊗ 核心放大: {otimes_count} 个（{int(OTIMES_CORE_HALF_PT)//2}pt, {font}）")

    total = oslash_count + bullet_count + large_count + otimes_count
    print(f"  QCA符号处理完成，共 {total} 个")
    return total


def _process_symbol_replace(root, w_ns, parent_map, src_char, dst_char,
                            font_size=None, font_name=None,
                            skip_if_has_size=False):
    """通用符号处理：替换字符（可选）、设置字号（可选）、设置字体（可选）

    Args:
        root: 文档根元素
        w_ns: Word命名空间URI
        parent_map: {id(child): parent} 映射
        src_char: 源字符
        dst_char: 目标字符（与src_char相同则不替换）
        font_size: 字号（半磅单位字符串），None则不改变
        font_name: 字体名称，None则不改变
        skip_if_has_size: 若True，run已有字号设置时跳过字号（仅设字体）

    Returns:
        int: 处理的符号数量
    """
    processed_count = 0

    # 收集所有需要处理的run（避免在迭代过程中修改树结构）
    runs_to_process = []
    for run in root.iter(f'{{{w_ns}}}r'):
        t_elem = run.find(f'{{{w_ns}}}t')
        if t_elem is None or t_elem.text is None:
            continue
        if src_char in t_elem.text:
            runs_to_process.append(run)

    for run in runs_to_process:
        t_elem = run.find(f'{{{w_ns}}}t')
        text = t_elem.text
        count_in_run = text.count(src_char)

        # 检查run中是否只有目标符号（无其他文字）
        text_without_symbol = text.replace(src_char, '')
        if text_without_symbol == '':
            # 简单情况：run中只有目标符号
            if src_char != dst_char:
                t_elem.text = text.replace(src_char, dst_char)
            effective_size = font_size
            if skip_if_has_size and font_size is not None:
                rPr = run.find(f'{{{w_ns}}}rPr')
                if rPr is not None and rPr.find(f'{{{w_ns}}}sz') is not None:
                    effective_size = None
            _set_run_style(run, w_ns, font_size=effective_size, font_name=font_name)
            processed_count += count_in_run
        else:
            # 复杂情况：符号与其他文字混合，需要拆分run
            parent = parent_map.get(id(run))
            if parent is None:
                print(f"  警告：无法找到run的父元素，跳过: {text[:30]}")
                continue
            split_count = _split_and_shrink_run(
                parent, run, w_ns,
                src_char=src_char,
                dst_char=dst_char,
                font_size=font_size,
                font_name=font_name,
                skip_if_has_size=skip_if_has_size
            )
            processed_count += split_count

    return processed_count


def _set_run_style(run, w_ns, font_size=None, font_name=None):
    """设置run的字号和/或字体

    Args:
        run: w:r 元素
        w_ns: Word命名空间URI
        font_size: 字号（半磅单位的字符串），None则不改变
        font_name: 字体名称，None则不改变
    """
    if font_size is None and font_name is None:
        return

    rPr = run.find(f'{{{w_ns}}}rPr')
    if rPr is None:
        rPr = ET.Element(f'{{{w_ns}}}rPr')
        run.insert(0, rPr)

    if font_name is not None:
        rFonts = rPr.find(f'{{{w_ns}}}rFonts')
        if rFonts is None:
            rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
        rFonts.set(f'{{{w_ns}}}ascii', font_name)
        rFonts.set(f'{{{w_ns}}}hAnsi', font_name)
        rFonts.set(f'{{{w_ns}}}eastAsia', font_name)
        # 移除hint属性，避免Word按eastAsia映射覆盖字体选择
        hint_key = f'{{{w_ns}}}hint'
        if hint_key in rFonts.attrib:
            del rFonts.attrib[hint_key]

    if font_size is not None:
        # 设置 w:sz（西文字号）
        sz = rPr.find(f'{{{w_ns}}}sz')
        if sz is None:
            sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
        sz.set(f'{{{w_ns}}}val', font_size)

        # 设置 w:szCs（复杂脚本字号，保持一致）
        szCs = rPr.find(f'{{{w_ns}}}szCs')
        if szCs is None:
            szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
        szCs.set(f'{{{w_ns}}}val', font_size)


def _split_and_shrink_run(parent, run, w_ns, src_char, dst_char,
                          font_size=None, font_name=None,
                          skip_if_has_size=False):
    """拆分含目标符号的混合文本run，仅对符号部分设置样式

    将 "文字A{src}文字B{src}文字C" 拆分为：
    - run("文字A")   — 保持原rPr
    - run("{dst}")   — 原rPr + 字号/字体
    - run("文字B")   — 保持原rPr
    - run("{dst}")   — 原rPr + 字号/字体
    - run("文字C")   — 保持原rPr

    Args:
        parent: run的父元素（通常是w:p段落）
        run: 原始的w:r元素
        w_ns: Word命名空间URI
        src_char: 源字符
        dst_char: 目标字符
        font_size: 字号（半磅单位字符串），None则不改变
        font_name: 字体名称，None则不改变
        skip_if_has_size: 若True，sym_run已有字号时跳过字号设置

    Returns:
        int: 处理的符号数量
    """
    t_elem = run.find(f'{{{w_ns}}}t')
    text = t_elem.text
    symbol_count = text.count(src_char)

    # 获取原run的rPr（用于复制给新run）
    orig_rPr = run.find(f'{{{w_ns}}}rPr')

    # 按源字符分割文本
    segments = text.split(src_char)

    # 找到原run在parent中的位置
    children = list(parent)
    run_index = None
    for i, child in enumerate(children):
        if child is run:
            run_index = i
            break

    if run_index is None:
        return 0

    # 移除原run
    parent.remove(run)

    # 在原位置插入新的run
    insert_pos = run_index
    for i, seg in enumerate(segments):
        # 插入普通文本段（如果非空）
        if seg:
            new_run = _create_run_with_rPr(w_ns, seg, orig_rPr)
            parent.insert(insert_pos, new_run)
            insert_pos += 1

        # 在非最后一个segment后面插入样式化的目标符号
        if i < len(segments) - 1:
            sym_run = _create_run_with_rPr(w_ns, dst_char, orig_rPr)
            effective_size = font_size
            if skip_if_has_size and font_size is not None:
                sym_rPr = sym_run.find(f'{{{w_ns}}}rPr')
                if sym_rPr is not None and sym_rPr.find(f'{{{w_ns}}}sz') is not None:
                    effective_size = None
            _set_run_style(sym_run, w_ns, font_size=effective_size, font_name=font_name)
            parent.insert(insert_pos, sym_run)
            insert_pos += 1

    return symbol_count


def _create_run_with_rPr(w_ns, text, orig_rPr):
    """创建一个新的run，复制原始rPr属性

    Args:
        w_ns: Word命名空间URI
        text: 文本内容
        orig_rPr: 原始run的rPr元素（可为None）

    Returns:
        新的w:r元素
    """
    run = ET.Element(f'{{{w_ns}}}r')

    # 复制原rPr
    if orig_rPr is not None:
        new_rPr = copy.deepcopy(orig_rPr)
        run.append(new_rPr)

    # 创建文本元素
    t = ET.SubElement(run, f'{{{w_ns}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    return run
