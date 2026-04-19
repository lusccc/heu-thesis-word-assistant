"""
公式续接段落处理模块
处理紧跟在公式后面的段落，根据Markdown源文件中是否有空行判断是否需要续接
"""
import xml.etree.ElementTree as ET
import re
import os
from .utils import NAMESPACES


def remove_paragraph_spacing_and_indent(para):
    """移除段落的首行缩进和段前间距
    
    Args:
        para: 段落元素
    """
    w_ns = NAMESPACES['w']
    
    # 获取或创建 pPr
    pPr = para.find('w:pPr', NAMESPACES)
    if pPr is None:
        pPr = ET.Element(f'{{{w_ns}}}pPr')
        para.insert(0, pPr)
    
    # 处理首行缩进
    ind = pPr.find('w:ind', NAMESPACES)
    if ind is None:
        ind = ET.Element(f'{{{w_ns}}}ind')
        pPr.append(ind)
    
    # 设置首行缩进为0
    ind.set(f'{{{w_ns}}}firstLine', '0')
    ind.set(f'{{{w_ns}}}firstLineChars', '0')
    
    # 处理段落间距
    spacing = pPr.find('w:spacing', NAMESPACES)
    if spacing is None:
        spacing = ET.Element(f'{{{w_ns}}}spacing')
        pPr.append(spacing)
    
    # 设置段前间距为0
    spacing.set(f'{{{w_ns}}}before', '0')
    spacing.set(f'{{{w_ns}}}beforeLines', '0')
    spacing.set(f'{{{w_ns}}}beforeAutospacing', '0')


def get_paragraph_text(para):
    """获取段落的纯文本内容
    
    Args:
        para: 段落元素
        
    Returns:
        str: 段落文本
    """
    text_elements = para.findall('.//w:t', NAMESPACES)
    return ''.join([t.text for t in text_elements if t.text]).strip()


def is_equation_paragraph(para):
    """判断段落是否为公式段落
    
    Args:
        para: 段落元素
        
    Returns:
        bool: 是否为公式段落
    """
    math_para = para.find('.//m:oMathPara', NAMESPACES)
    return math_para is not None


def get_equation_id(para):
    """获取公式段落的ID（书签名）
    
    Args:
        para: 段落元素
        
    Returns:
        str: 公式ID，如'eq-attention'，如果没有则返回None
    """
    for bookmark in para.findall('.//w:bookmarkStart', NAMESPACES):
        name = bookmark.get(f'{{{NAMESPACES["w"]}}}name')
        if name and name.startswith('eq-'):
            return name
    return None


def parse_qmd_equations_without_blank_lines(qmd_path):
    """解析QMD文件，找出公式后没有空行的公式ID
    
    Args:
        qmd_path: QMD文件路径
        
    Returns:
        set: 公式后没有空行的公式ID集合
    """
    if not os.path.exists(qmd_path):
        print(f"  警告：找不到QMD文件 {qmd_path}，将使用默认处理方式")
        return set()
    
    try:
        with open(qmd_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"  警告：读取QMD文件失败 {e}，将使用默认处理方式")
        return set()
    
    equations_without_blank = set()
    
    # 正则表达式匹配公式标签行：$$ {#eq-xxx}
    eq_label_pattern = re.compile(r'^\$\$\s*\{#(eq-[^}]+)\}\s*$')
    
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        
        # 检查是否是公式标签行
        match = eq_label_pattern.match(line)
        if match:
            eq_id = match.group(1)
            
            # 检查下一行是否为空行
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if next_line:  # 下一行不是空行
                    equations_without_blank.add(eq_id)
        
        i += 1
    
    return equations_without_blank


def process_equation_continuation_paragraphs(root, qmd_path=None):
    """处理紧跟在公式后面的段落
    
    根据Markdown源文件中公式后是否有空行来判断：
    - 如果公式后没有空行，下一段应该紧跟公式（移除缩进和间距）
    - 如果公式后有空行，下一段是新段落（保持正常格式）
    
    Args:
        root: 文档根元素
        qmd_path: QMD源文件路径（可选）
        
    Returns:
        int: 处理的段落数量
    """
    modified_count = 0
    
    print("正在处理公式后的续接段落...")
    
    # 解析QMD文件，获取公式后没有空行的公式ID集合
    equations_without_blank = set()
    if qmd_path and os.path.exists(qmd_path):
        equations_without_blank = parse_qmd_equations_without_blank_lines(qmd_path)
        print(f"  从QMD文件中识别出 {len(equations_without_blank)} 个公式后没有空行")
    else:
        print("  未提供QMD文件路径，将使用默认处理方式")
    
    # 获取所有段落
    all_paras = root.findall('.//w:p', NAMESPACES)
    
    for i in range(len(all_paras) - 1):
        current_para = all_paras[i]
        next_para = all_paras[i + 1]
        
        # 检查当前段落是否为公式段落
        if is_equation_paragraph(current_para):
            # 获取公式ID
            eq_id = get_equation_id(current_para)
            
            # 如果这个公式后没有空行，就处理下一段
            if eq_id and eq_id in equations_without_blank:
                # 移除首行缩进和段前间距
                remove_paragraph_spacing_and_indent(next_para)
                modified_count += 1
                next_text = get_paragraph_text(next_para)
                print(f"  处理公式 {eq_id} 后的续接段落: {next_text[:30]}...")
    
    print(f"  已处理 {modified_count} 个公式后的续接段落")
    return modified_count
