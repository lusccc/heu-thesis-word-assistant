"""
章节标题样式处理模块
功能：移除非正文章节（结论、参考文献、致谢等）的编号
"""
import xml.etree.ElementTree as ET
from .utils import NAMESPACES
from .config import EXCLUDE_CHAPTER_TITLES, STYLES

# 需要在两字之间保留四个空格的标题
TWO_CHAR_TITLES_WITH_SPACES = ['摘要', '结论', '致谢']


def remove_numbering_from_excluded_headings(root):
    """移除非正文章节标题的编号引用
    
    模板中一级标题样式关联了编号列表（numId: 3），会自动添加"第X章"。
    对于结论、参考文献、致谢等非正文章节，需要通过设置 numId=0 来禁用编号。
    """
    print("正在处理章节标题编号...")
    
    w_ns = NAMESPACES['w']
    count = 0
    
    for para in root.findall(f'.//{{{w_ns}}}p'):
        pPr = para.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue
        
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue
        
        style_val = pStyle.get(f'{{{w_ns}}}val', '')
        # 检查是否是一级标题
        if style_val not in (STYLES['heading1'], STYLES['heading1_alt']):
            continue
        
        # 提取标题文本
        title_text = ''
        for t in para.findall(f'.//{{{w_ns}}}t'):
            if t.text:
                title_text += t.text
        title_text = title_text.strip()
        
        # 检查是否是需要排除编号的章节
        # 去除空格后比较，因为Quarto会将多空格压缩为单空格（如"摘    要"→"摘 要"）
        title_text_nospace = title_text.replace(' ', '').replace('\u3000', '')
        is_excluded = any(exclude in title_text_nospace for exclude in EXCLUDE_CHAPTER_TITLES)
        
        if is_excluded:
            # 查找或创建 numPr
            numPr = pPr.find(f'{{{w_ns}}}numPr')
            if numPr is None:
                numPr = ET.Element(f'{{{w_ns}}}numPr')
                # 插入到 pStyle 之后
                pStyle_index = list(pPr).index(pStyle)
                pPr.insert(pStyle_index + 1, numPr)
            
            # 设置 numId=0 来禁用编号
            numId = numPr.find(f'{{{w_ns}}}numId')
            if numId is None:
                numId = ET.SubElement(numPr, f'{{{w_ns}}}numId')
            numId.set(f'{{{w_ns}}}val', '0')
            
            # 设置 ilvl=0
            ilvl = numPr.find(f'{{{w_ns}}}ilvl')
            if ilvl is None:
                ilvl = ET.SubElement(numPr, f'{{{w_ns}}}ilvl')
            ilvl.set(f'{{{w_ns}}}val', '0')
            
            count += 1
            
            # 对两字标题修复空格：将Quarto压缩的单空格恢复为四个空格
            if title_text_nospace in TWO_CHAR_TITLES_WITH_SPACES:
                for run in para.findall(f'{{{w_ns}}}r'):
                    t_elem = run.find(f'{{{w_ns}}}t')
                    if t_elem is not None and t_elem.text and t_elem.text.strip() == '':
                        t_elem.text = '    '
                        print(f"  修复空格: {title_text_nospace} (1→4个空格)")
            
            print(f"  禁用编号: {title_text}")
    
    print(f"  共处理 {count} 个非正文章节")
    return count
