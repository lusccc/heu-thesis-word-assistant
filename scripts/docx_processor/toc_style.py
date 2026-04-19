"""
目录样式处理模块
功能：
1. 将目录标题改为"目    录"（使用"标题1"样式）
2. 在目录后添加分节符（nextPage），使第一章从下一页开始
3. 目录部分不显示页码（页眉由header_footer.py处理）
4. 第一章页码从1开始
5. 使用下一页分节符，避免强制插入空白页
"""
import xml.etree.ElementTree as ET

# 命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

# 目录标题文本（中间有4个全角空格）
TOC_TITLE = "目    录"


def process_toc(root):
    """处理目录样式
    
    1. 修改目录标题为"目    录"，使用"标题1"样式
    2. 在目录后添加分节符（nextPage），使第一章从下一页开始且页码从1开始
    3. 目录部分不显示页码
    """
    print("正在处理目录样式...")
    
    w_ns = NAMESPACES['w']
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        raise ValueError("找不到document body")
    
    # 1. 找到目录的sdt元素
    sdt = body.find(f'{{{w_ns}}}sdt')
    if sdt is None:
        print("  未找到目录元素，跳过目录处理")
        return
    
    sdt_content = sdt.find(f'{{{w_ns}}}sdtContent')
    if sdt_content is None:
        print("  未找到目录内容，跳过目录处理")
        return
    
    # 2. 修改目录标题
    _modify_toc_title(sdt_content, w_ns)
    
    # 3. 在目录后添加分节符（在sdt元素后面）
    _add_section_break_after_toc(body, sdt, w_ns)
    
    print("  目录样式处理完成")


def _modify_toc_title(sdt_content, w_ns):
    """修改目录标题为"目    录"，应用标题1的视觉格式但不被TOC收录
    
    不使用pStyle引用标题1样式（会被TOC收录），而是直接应用格式属性
    """
    # 找到第一个段落（目录标题）
    first_para = sdt_content.find(f'{{{w_ns}}}p')
    if first_para is None:
        return
    
    # 获取或创建pPr
    pPr = first_para.find(f'{{{w_ns}}}pPr')
    if pPr is None:
        pPr = ET.Element(f'{{{w_ns}}}pPr')
        first_para.insert(0, pPr)
    
    # 设置pStyle为Heading1（使VBA宏能识别目录标题）
    pStyle = pPr.find(f'{{{w_ns}}}pStyle')
    if pStyle is None:
        pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
    pStyle.set(f'{{{w_ns}}}val', '1')
    
    # 移除numPr（编号）
    numPr = pPr.find(f'{{{w_ns}}}numPr')
    if numPr is not None:
        pPr.remove(numPr)
    
    # 设置outlineLvl=9，使目录标题不被TOC收录（_TocRange书签也提供了保护）
    outlineLvl = pPr.find(f'{{{w_ns}}}outlineLvl')
    if outlineLvl is None:
        outlineLvl = ET.SubElement(pPr, f'{{{w_ns}}}outlineLvl')
    outlineLvl.set(f'{{{w_ns}}}val', '9')
    
    # 应用标题1的段落格式
    # keepNext
    if pPr.find(f'{{{w_ns}}}keepNext') is None:
        ET.SubElement(pPr, f'{{{w_ns}}}keepNext')
    # keepLines
    if pPr.find(f'{{{w_ns}}}keepLines') is None:
        ET.SubElement(pPr, f'{{{w_ns}}}keepLines')
    # spacing
    spacing = pPr.find(f'{{{w_ns}}}spacing')
    if spacing is None:
        spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
    spacing.set(f'{{{w_ns}}}beforeLines', '100')
    spacing.set(f'{{{w_ns}}}before', '100')
    spacing.set(f'{{{w_ns}}}afterLines', '100')
    spacing.set(f'{{{w_ns}}}after', '100')
    spacing.set(f'{{{w_ns}}}line', '440')
    spacing.set(f'{{{w_ns}}}lineRule', 'exact')
    # 移除首行缩进，避免继承正文样式的缩进
    ind = pPr.find(f'{{{w_ns}}}ind')
    if ind is None:
        ind = ET.SubElement(pPr, f'{{{w_ns}}}ind')
    ind.set(f'{{{w_ns}}}firstLine', '0')
    ind.set(f'{{{w_ns}}}firstLineChars', '0')
    ind.set(f'{{{w_ns}}}left', '0')
    ind.set(f'{{{w_ns}}}leftChars', '0')
    # jc (居中)
    jc = pPr.find(f'{{{w_ns}}}jc')
    if jc is None:
        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
    jc.set(f'{{{w_ns}}}val', 'center')
    
    # 修改文本内容和字符格式
    for r in first_para.findall(f'{{{w_ns}}}r'):
        t = r.find(f'{{{w_ns}}}t')
        if t is not None:
            t.text = TOC_TITLE
            
            # 获取或创建rPr
            rPr = r.find(f'{{{w_ns}}}rPr')
            if rPr is None:
                rPr = ET.Element(f'{{{w_ns}}}rPr')
                r.insert(0, rPr)
            
            # 应用标题1的字符格式
            # rFonts
            rFonts = rPr.find(f'{{{w_ns}}}rFonts')
            if rFonts is None:
                rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
            rFonts.set(f'{{{w_ns}}}ascii', 'Times New Roman')
            rFonts.set(f'{{{w_ns}}}eastAsia', '黑体')
            rFonts.set(f'{{{w_ns}}}hAnsi', 'Times New Roman')
            rFonts.set(f'{{{w_ns}}}hint', 'eastAsia')
            # sz (字号18磅=36半磅)
            sz = rPr.find(f'{{{w_ns}}}sz')
            if sz is None:
                sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
            sz.set(f'{{{w_ns}}}val', '36')
            szCs = rPr.find(f'{{{w_ns}}}szCs')
            if szCs is None:
                szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
            szCs.set(f'{{{w_ns}}}val', '36')
    
    print(f"  已将目录标题修改为：{TOC_TITLE}（不被TOC收录）")


def _add_section_break_after_toc(body, sdt, w_ns):
    """在目录后添加分节符
    
    分节符设置：
    - type: nextPage（从下一页开始，不强制奇数页）
    - 目录部分不显示页码（通过不添加页脚引用实现）
    
    注意：此sectPr定义的是目录部分的属性，不是第1章的属性！
    第1章的页码从1开始由header_footer.py中设置。
    """
    from .page_style import PAGE_WIDTH, PAGE_HEIGHT
    
    # 找到sdt在body中的位置
    children = list(body)
    sdt_index = children.index(sdt)
    
    # 创建包含分节符的段落
    sect_para = ET.Element(f'{{{w_ns}}}p')
    pPr = ET.SubElement(sect_para, f'{{{w_ns}}}pPr')
    sectPr = ET.SubElement(pPr, f'{{{w_ns}}}sectPr')
    
    # 设置分节符类型为nextPage（使下一节从下一页开始）
    sect_type = ET.SubElement(sectPr, f'{{{w_ns}}}type')
    sect_type.set(f'{{{w_ns}}}val', 'nextPage')
    
    # 添加页面大小（A4）
    pgSz = ET.SubElement(sectPr, f'{{{w_ns}}}pgSz')
    pgSz.set(f'{{{w_ns}}}w', PAGE_WIDTH)
    pgSz.set(f'{{{w_ns}}}h', PAGE_HEIGHT)
    
    # 添加页边距（从config获取）
    from .config import get_margin_twips
    pgMar = ET.SubElement(sectPr, f'{{{w_ns}}}pgMar')
    pgMar.set(f'{{{w_ns}}}top', get_margin_twips('top'))
    pgMar.set(f'{{{w_ns}}}bottom', get_margin_twips('bottom'))
    pgMar.set(f'{{{w_ns}}}left', get_margin_twips('left'))
    pgMar.set(f'{{{w_ns}}}right', get_margin_twips('right'))
    pgMar.set(f'{{{w_ns}}}header', get_margin_twips('header'))
    pgMar.set(f'{{{w_ns}}}footer', get_margin_twips('footer'))
    pgMar.set(f'{{{w_ns}}}gutter', get_margin_twips('gutter'))
    
    # 在sdt后面插入分节符段落
    body.insert(sdt_index + 1, sect_para)
    
    print("  已在目录后添加分节符（nextPage，使第1章从下一页开始）")
