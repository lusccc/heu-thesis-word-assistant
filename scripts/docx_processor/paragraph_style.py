"""
段落样式处理模块
处理含公式段落的行间距等样式
"""
import xml.etree.ElementTree as ET
from .utils import NAMESPACES

# 行间距配置（atLeast 模式）
# 最小行距值与模板 Normal 样式一致（440 twips = 22pt）
# 不含公式的行保持正常行距，含公式的行自动撑开
LINE_SPACING_VALUE = 440  # twips (1/20 pt)
LINE_SPACING_RULE = 'atLeast'


def set_paragraph_line_spacing(para, spacing_value=LINE_SPACING_VALUE):
    """设置段落行间距为 atLeast 模式
    
    Args:
        para: 段落元素
        spacing_value: 最小行间距值（单位：twips，1/20 磅）
    """
    w_ns = NAMESPACES['w']
    
    # 获取或创建 pPr
    pPr = para.find('w:pPr', NAMESPACES)
    if pPr is None:
        pPr = ET.Element(f'{{{w_ns}}}pPr')
        para.insert(0, pPr)
    
    # 获取或创建 spacing
    spacing = pPr.find('w:spacing', NAMESPACES)
    if spacing is None:
        spacing = ET.Element(f'{{{w_ns}}}spacing')
        pPr.append(spacing)
    
    # 设置行间距
    # line: 最小行间距值（twips）
    # lineRule: "atLeast" 表示最小行距，含公式的行自动撑开
    spacing.set(f'{{{w_ns}}}line', str(spacing_value))
    spacing.set(f'{{{w_ns}}}lineRule', LINE_SPACING_RULE)


def process_paragraphs_with_math(root):
    """处理含有行内公式的段落，设置行间距为 atLeast 模式
    
    Args:
        root: 文档根元素
    
    Returns:
        int: 处理的段落数量
    """
    modified_count = 0
    w_ns = NAMESPACES['w']
    m_ns = NAMESPACES['m']
    
    print("正在处理含公式段落的行间距...")
    
    # 收集表格内的段落，跳过它们（表格有自己的行距样式）
    table_paras = set()
    for tbl in root.iter(f'{{{w_ns}}}tbl'):
        for p in tbl.iter(f'{{{w_ns}}}p'):
            table_paras.add(p)
    
    for para in root.findall('.//w:p', NAMESPACES):
        if para in table_paras:
            continue
        # 检查段落是否包含行内公式（m:oMath，但不是 m:oMathPara）
        # m:oMathPara 是独立的公式段落，不需要调整行间距
        has_inline_math = False
        
        # 查找行内公式
        inline_math = para.findall('.//m:oMath', NAMESPACES)
        math_para = para.find('.//m:oMathPara', NAMESPACES)
        
        # 如果有 oMath 但没有 oMathPara，说明是行内公式
        if inline_math and math_para is None:
            has_inline_math = True
        
        if has_inline_math:
            set_paragraph_line_spacing(para)
            modified_count += 1
    
    print(f"  已为 {modified_count} 个含行内公式的段落设置 atLeast {LINE_SPACING_VALUE} twips 行间距")
    return modified_count


def strip_spaces_around_inline_math(root):
    """全局去除行内公式(m:oMath)前后的多余空格run
    
    Pandoc 将 Markdown 中 $...$ 前后的空格生成为独立的 <w:r> 空格 run，
    导致 Word 中公式前后出现多余间距。本函数移除这些空格 run 或
    截断相邻 run 尾部/头部的空格。
    """
    w_ns = NAMESPACES['w']
    m_ns = NAMESPACES['m']
    
    print("正在去除行内公式前后的多余空格...")
    removed_count = 0
    trimmed_count = 0
    
    for para in root.iter(f'{{{w_ns}}}p'):
        children = list(para)
        # 快速检查是否包含 oMath
        if not any(c.tag == f'{{{m_ns}}}oMath' for c in children):
            continue
        
        to_remove = set()
        
        for idx, child in enumerate(children):
            if child.tag != f'{{{m_ns}}}oMath':
                continue
            
            # 处理 oMath 前面的 run
            if idx > 0 and children[idx - 1].tag == f'{{{w_ns}}}r':
                prev_run = children[idx - 1]
                prev_t = prev_run.find(f'{{{w_ns}}}t')
                if prev_t is not None and prev_t.text is not None:
                    if prev_t.text.strip() == '':
                        # 纯空格 run，标记移除
                        to_remove.add(id(prev_run))
                        removed_count += 1
                    elif prev_t.text.endswith(' '):
                        prev_t.text = prev_t.text.rstrip(' ')
                        trimmed_count += 1
            
            # 处理 oMath 后面的 run
            if idx < len(children) - 1 and children[idx + 1].tag == f'{{{w_ns}}}r':
                next_run = children[idx + 1]
                next_t = next_run.find(f'{{{w_ns}}}t')
                if next_t is not None and next_t.text is not None:
                    if next_t.text.strip() == '':
                        # 纯空格 run，标记移除
                        to_remove.add(id(next_run))
                        removed_count += 1
                    elif next_t.text.startswith(' '):
                        next_t.text = next_t.text.lstrip(' ')
                        trimmed_count += 1
        
        # 移除标记的 run
        for child in children:
            if id(child) in to_remove:
                para.remove(child)
    
    print(f"  移除 {removed_count} 个纯空格run，修剪 {trimmed_count} 个run的空格")
