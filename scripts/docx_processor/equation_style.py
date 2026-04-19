import xml.etree.ElementTree as ET
from .utils import NAMESPACES
from .config import MATH_FONT, EXCLUDE_CHAPTER_TITLES, STYLES, EQUATION_TAB_CENTER, EQUATION_TAB_RIGHT


def fix_bold_math_to_bold_italic(root):
    """将OMML中的粗体数学样式 b 修正为粗斜体 bi。"""
    m_ns = NAMESPACES['m']
    fixed_count = 0

    for sty in root.findall('.//m:rPr/m:sty', NAMESPACES):
        val = sty.get(f'{{{m_ns}}}val')
        if val == 'b':
            sty.set(f'{{{m_ns}}}val', 'bi')
            fixed_count += 1

    if fixed_count > 0:
        print(f"已将 {fixed_count} 处数学样式从 b 修正为 bi")

    return fixed_count

def _build_number_elements(m_ns, w_ns, chapter_num, eq_num):
    """构建公式编号的所有元素列表：（STYLEREF-SEQ）"""
    elements = []
    
    # 左括号（
    lparen_run = ET.Element(f'{{{m_ns}}}r')
    lparen_rPr = ET.Element(f'{{{m_ns}}}rPr')
    lparen_nor = ET.Element(f'{{{m_ns}}}nor')
    lparen_rPr.append(lparen_nor)
    lparen_run.append(lparen_rPr)
    lparen_w_rPr = ET.Element(f'{{{w_ns}}}rPr')
    lparen_rFonts = ET.Element(f'{{{w_ns}}}rFonts')
    lparen_rFonts.set(f'{{{w_ns}}}hint', 'eastAsia')
    lparen_w_rPr.append(lparen_rFonts)
    lparen_run.append(lparen_w_rPr)
    lparen_t = ET.Element(f'{{{m_ns}}}t')
    lparen_t.text = '（'
    lparen_run.append(lparen_t)
    elements.append(lparen_run)
    
    # STYLEREF域 - begin
    styleref_begin = ET.Element(f'{{{m_ns}}}r')
    styleref_begin_rPr = ET.Element(f'{{{m_ns}}}rPr')
    styleref_begin_nor = ET.Element(f'{{{m_ns}}}nor')
    styleref_begin_rPr.append(styleref_begin_nor)
    styleref_begin.append(styleref_begin_rPr)
    styleref_begin_fld = ET.Element(f'{{{w_ns}}}fldChar')
    styleref_begin_fld.set(f'{{{w_ns}}}fldCharType', 'begin')
    styleref_begin_fld.set(f'{{{w_ns}}}fldLock', '1')
    styleref_begin.append(styleref_begin_fld)
    elements.append(styleref_begin)
    
    # STYLEREF域 - code
    styleref_code = ET.Element(f'{{{m_ns}}}r')
    styleref_code_rPr = ET.Element(f'{{{m_ns}}}rPr')
    styleref_code_nor = ET.Element(f'{{{m_ns}}}nor')
    styleref_code_rPr.append(styleref_code_nor)
    styleref_code.append(styleref_code_rPr)
    styleref_code_t = ET.Element(f'{{{m_ns}}}t')
    styleref_code_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    styleref_code_t.text = ' STYLEREF 1 \\s '
    styleref_code.append(styleref_code_t)
    elements.append(styleref_code)
    
    # STYLEREF域 - separate
    styleref_sep = ET.Element(f'{{{m_ns}}}r')
    styleref_sep_rPr = ET.Element(f'{{{m_ns}}}rPr')
    styleref_sep_nor = ET.Element(f'{{{m_ns}}}nor')
    styleref_sep_rPr.append(styleref_sep_nor)
    styleref_sep.append(styleref_sep_rPr)
    styleref_sep_fld = ET.Element(f'{{{w_ns}}}fldChar')
    styleref_sep_fld.set(f'{{{w_ns}}}fldCharType', 'separate')
    styleref_sep.append(styleref_sep_fld)
    elements.append(styleref_sep)
    
    # STYLEREF域 - display
    styleref_disp = ET.Element(f'{{{m_ns}}}r')
    styleref_disp_rPr = ET.Element(f'{{{m_ns}}}rPr')
    styleref_disp_nor = ET.Element(f'{{{m_ns}}}nor')
    styleref_disp_rPr.append(styleref_disp_nor)
    styleref_disp.append(styleref_disp_rPr)
    styleref_disp_t = ET.Element(f'{{{m_ns}}}t')
    styleref_disp_t.text = str(chapter_num)
    styleref_disp.append(styleref_disp_t)
    elements.append(styleref_disp)
    
    # STYLEREF域 - end
    styleref_end = ET.Element(f'{{{m_ns}}}r')
    styleref_end_rPr = ET.Element(f'{{{m_ns}}}rPr')
    styleref_end_nor = ET.Element(f'{{{m_ns}}}nor')
    styleref_end_rPr.append(styleref_end_nor)
    styleref_end.append(styleref_end_rPr)
    styleref_end_fld = ET.Element(f'{{{w_ns}}}fldChar')
    styleref_end_fld.set(f'{{{w_ns}}}fldCharType', 'end')
    styleref_end.append(styleref_end_fld)
    elements.append(styleref_end)
    
    # 连字符 -
    dash_run = ET.Element(f'{{{m_ns}}}r')
    dash_rPr = ET.Element(f'{{{m_ns}}}rPr')
    dash_nor = ET.Element(f'{{{m_ns}}}nor')
    dash_rPr.append(dash_nor)
    dash_run.append(dash_rPr)
    dash_t = ET.Element(f'{{{m_ns}}}t')
    dash_t.text = '-'
    dash_run.append(dash_t)
    elements.append(dash_run)
    
    # SEQ域 - begin
    seq_begin = ET.Element(f'{{{m_ns}}}r')
    seq_begin_rPr = ET.Element(f'{{{m_ns}}}rPr')
    seq_begin_nor = ET.Element(f'{{{m_ns}}}nor')
    seq_begin_rPr.append(seq_begin_nor)
    seq_begin.append(seq_begin_rPr)
    seq_begin_fld = ET.Element(f'{{{w_ns}}}fldChar')
    seq_begin_fld.set(f'{{{w_ns}}}fldCharType', 'begin')
    seq_begin_fld.set(f'{{{w_ns}}}fldLock', '1')
    seq_begin.append(seq_begin_fld)
    elements.append(seq_begin)
    
    # SEQ域 - code
    seq_code = ET.Element(f'{{{m_ns}}}r')
    seq_code_rPr = ET.Element(f'{{{m_ns}}}rPr')
    seq_code_nor = ET.Element(f'{{{m_ns}}}nor')
    seq_code_rPr.append(seq_code_nor)
    seq_code.append(seq_code_rPr)
    seq_code_t = ET.Element(f'{{{m_ns}}}t')
    seq_code_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    seq_code_t.text = ' SEQ Equation \\* ARABIC \\s 1 '
    seq_code.append(seq_code_t)
    elements.append(seq_code)
    
    # SEQ域 - separate
    seq_sep = ET.Element(f'{{{m_ns}}}r')
    seq_sep_rPr = ET.Element(f'{{{m_ns}}}rPr')
    seq_sep_nor = ET.Element(f'{{{m_ns}}}nor')
    seq_sep_rPr.append(seq_sep_nor)
    seq_sep.append(seq_sep_rPr)
    seq_sep_fld = ET.Element(f'{{{w_ns}}}fldChar')
    seq_sep_fld.set(f'{{{w_ns}}}fldCharType', 'separate')
    seq_sep.append(seq_sep_fld)
    elements.append(seq_sep)
    
    # SEQ域 - display
    seq_disp = ET.Element(f'{{{m_ns}}}r')
    seq_disp_rPr = ET.Element(f'{{{m_ns}}}rPr')
    seq_disp_nor = ET.Element(f'{{{m_ns}}}nor')
    seq_disp_rPr.append(seq_disp_nor)
    seq_disp.append(seq_disp_rPr)
    seq_disp_t = ET.Element(f'{{{m_ns}}}t')
    seq_disp_t.text = str(eq_num)
    seq_disp.append(seq_disp_t)
    elements.append(seq_disp)
    
    # SEQ域 - end
    seq_end = ET.Element(f'{{{m_ns}}}r')
    seq_end_rPr = ET.Element(f'{{{m_ns}}}rPr')
    seq_end_nor = ET.Element(f'{{{m_ns}}}nor')
    seq_end_rPr.append(seq_end_nor)
    seq_end.append(seq_end_rPr)
    seq_end_fld = ET.Element(f'{{{w_ns}}}fldChar')
    seq_end_fld.set(f'{{{w_ns}}}fldCharType', 'end')
    seq_end.append(seq_end_fld)
    elements.append(seq_end)
    
    # 右括号）
    rparen_run = ET.Element(f'{{{m_ns}}}r')
    rparen_rPr = ET.Element(f'{{{m_ns}}}rPr')
    rparen_nor = ET.Element(f'{{{m_ns}}}nor')
    rparen_rPr.append(rparen_nor)
    rparen_run.append(rparen_rPr)
    rparen_w_rPr = ET.Element(f'{{{w_ns}}}rPr')
    rparen_rFonts = ET.Element(f'{{{w_ns}}}rFonts')
    rparen_rFonts.set(f'{{{w_ns}}}hint', 'eastAsia')
    rparen_w_rPr.append(rparen_rFonts)
    rparen_run.append(rparen_w_rPr)
    rparen_t = ET.Element(f'{{{m_ns}}}t')
    rparen_t.text = '）'
    rparen_run.append(rparen_t)
    elements.append(rparen_run)
    
    return elements


def wrap_in_eqarr_and_add_number(omath, chapter_num, eq_num):
    """将oMath的内容包装在eqArr>e结构中，并添加#和编号。
    依赖 eqArr + maxDist=1 + # 机制实现公式居中、编号右对齐。"""
    m_ns = NAMESPACES['m']
    w_ns = NAMESPACES['w']
    
    # 保存原有的所有子元素
    original_children = list(omath)
    
    # 清理公式末尾的空白 m:r 元素（Quarto 渲染时可能添加的 EM QUAD 等空格）
    while original_children:
        last = original_children[-1]
        if last.tag == f'{{{m_ns}}}r':
            t_elem = last.find(f'{{{m_ns}}}t')
            if t_elem is not None and t_elem.text and t_elem.text.strip() == '':
                original_children.pop()
                continue
        break
    
    # 清空oMath
    for child in list(omath):
        omath.remove(child)
    
    # 先构建编号部分的元素
    number_elements = _build_number_elements(m_ns, w_ns, chapter_num, eq_num)
    
    # 创建eqArr结构
    eqArr = ET.Element(f'{{{m_ns}}}eqArr')
    
    # 创建eqArrPr（属性）
    eqArrPr = ET.Element(f'{{{m_ns}}}eqArrPr')
    maxDist = ET.Element(f'{{{m_ns}}}maxDist')
    maxDist.set(f'{{{m_ns}}}val', '1')
    eqArrPr.append(maxDist)
    
    ctrlPr = ET.Element(f'{{{m_ns}}}ctrlPr')
    rPr = ET.Element(f'{{{w_ns}}}rPr')
    rFonts = ET.Element(f'{{{w_ns}}}rFonts')
    rFonts.set(f'{{{w_ns}}}ascii', MATH_FONT)
    rFonts.set(f'{{{w_ns}}}hAnsi', MATH_FONT)
    rPr.append(rFonts)
    i_elem = ET.Element(f'{{{w_ns}}}i')
    rPr.append(i_elem)
    ctrlPr.append(rPr)
    eqArrPr.append(ctrlPr)
    
    eqArr.append(eqArrPr)
    
    # 创建e元素（包含公式内容）
    e = ET.Element(f'{{{m_ns}}}e')
    
    # 将原有的公式内容添加到e元素中
    for child in original_children:
        e.append(child)
    
    # 在e元素中添加#号
    hash_run = ET.Element(f'{{{m_ns}}}r')
    hash_rPr = ET.Element(f'{{{w_ns}}}rPr')
    hash_rFonts = ET.Element(f'{{{w_ns}}}rFonts')
    hash_rFonts.set(f'{{{w_ns}}}ascii', MATH_FONT)
    hash_rFonts.set(f'{{{w_ns}}}hAnsi', MATH_FONT)
    hash_rPr.append(hash_rFonts)
    hash_run.append(hash_rPr)
    hash_t = ET.Element(f'{{{m_ns}}}t')
    hash_t.text = '#'
    hash_run.append(hash_t)
    e.append(hash_run)
    
    # 添加编号元素
    for ne in number_elements:
        e.append(ne)
    
    # 添加ctrlPr
    e_ctrlPr = ET.Element(f'{{{m_ns}}}ctrlPr')
    e_rPr = ET.Element(f'{{{w_ns}}}rPr')
    e_rFonts = ET.Element(f'{{{w_ns}}}rFonts')
    e_rFonts.set(f'{{{w_ns}}}ascii', MATH_FONT)
    e_rFonts.set(f'{{{w_ns}}}hAnsi', MATH_FONT)
    e_rPr.append(e_rFonts)
    e_ctrlPr.append(e_rPr)
    e.append(e_ctrlPr)
    
    # 将e添加到eqArr
    eqArr.append(e)
    
    # 将eqArr添加到oMath
    omath.append(eqArr)

def set_equation_style(para):
    """设置段落为公式样式，并添加制表位用于公式居中+编号右对齐"""
    w_ns = NAMESPACES['w']
    pPr = para.find('w:pPr', NAMESPACES)
    if pPr is None:
        pPr = ET.Element(f'{{{w_ns}}}pPr')
        para.insert(0, pPr)
    
    pStyle = pPr.find('w:pStyle', NAMESPACES)
    if pStyle is None:
        pStyle = ET.Element(f'{{{w_ns}}}pStyle')
        pPr.insert(0, pStyle)
    
    pStyle.set(f'{{{w_ns}}}val', STYLES['equation'])
    
    # 显式设置段落居中对齐（不依赖样式继承）
    jc = pPr.find(f'{{{w_ns}}}jc')
    if jc is None:
        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
    jc.set(f'{{{w_ns}}}val', 'center')
    
    # 显式设置零缩进，防止继承正文样式的首行缩进等
    ind = pPr.find(f'{{{w_ns}}}ind')
    if ind is None:
        ind = ET.SubElement(pPr, f'{{{w_ns}}}ind')
    ind.set(f'{{{w_ns}}}firstLine', '0')
    ind.set(f'{{{w_ns}}}left', '0')
    ind.set(f'{{{w_ns}}}right', '0')
    
    # 添加制表位：居中（公式位置）和右对齐（编号位置）
    # 配合 m:eqArr + # 分隔符实现公式居中、编号右对齐
    old_tabs = pPr.find(f'{{{w_ns}}}tabs')
    if old_tabs is not None:
        pPr.remove(old_tabs)
    tabs = ET.SubElement(pPr, f'{{{w_ns}}}tabs')
    tab_c = ET.SubElement(tabs, f'{{{w_ns}}}tab')
    tab_c.set(f'{{{w_ns}}}val', 'center')
    tab_c.set(f'{{{w_ns}}}pos', EQUATION_TAB_CENTER)
    tab_r = ET.SubElement(tabs, f'{{{w_ns}}}tab')
    tab_r.set(f'{{{w_ns}}}val', 'right')
    tab_r.set(f'{{{w_ns}}}pos', EQUATION_TAB_RIGHT)

def process_equations(root):
    """处理文档中的公式"""
    fix_bold_math_to_bold_italic(root)

    modified_count = 0
    equation_number = 0
    current_chapter = 0
    
    # 存储公式ID到编号的映射
    equation_map = {}
    
    for para in root.findall('.//w:p', NAMESPACES):
        # 1. 章节检测
        pPr = para.find('w:pPr', NAMESPACES)
        if pPr is not None:
            pStyle = pPr.find('w:pStyle', NAMESPACES)
            if pStyle is not None:
                style_val = pStyle.get(f'{{{NAMESPACES["w"]}}}val')
                if style_val == STYLES['heading1']:
                    text_elements = para.findall('.//w:t', NAMESPACES)
                    para_text = ''.join([t.text for t in text_elements if t.text]).strip()
                    
                    para_text_nospace = para_text.replace(' ', '').replace('\u3000', '')
                    is_excluded = any(exclude in para_text_nospace for exclude in EXCLUDE_CHAPTER_TITLES)
                    
                    if not is_excluded:
                        current_chapter += 1
                        equation_number = 0
                        print(f"进入第 {current_chapter} 章: {para_text}")
                    else:
                        print(f"跳过非章节标题: {para_text}")
        
        # 2. 公式处理
        math_para = para.find('.//m:oMathPara', NAMESPACES)
        if math_para is not None:
            effective_chapter = current_chapter if current_chapter > 0 else 0
            equation_number += 1
            
            # 记录书签映射
            for bookmark in para.findall('.//w:bookmarkStart', NAMESPACES):
                name = bookmark.get(f'{{{NAMESPACES["w"]}}}name')
                if name and name.startswith('eq-'):
                    equation_map[name] = (effective_chapter, equation_number)
                    print(f"  映射公式 {name} -> ({effective_chapter}-{equation_number})")
            
            # 设置样式
            set_equation_style(para)
            
            # 处理oMath内容
            omath = math_para.find('.//m:oMath', NAMESPACES)
            if omath is not None:
                eqArr = omath.find('.//m:eqArr', NAMESPACES)
                if eqArr is None:
                    # 移除旧编号
                    delimiters = omath.findall('.//m:d', NAMESPACES)
                    if delimiters:
                        last_d = delimiters[-1]
                        texts = [t.text for t in last_d.findall('.//m:t', NAMESPACES) if t.text and t.text.strip()]
                        if len(texts) == 1 and texts[0].isdigit():
                            for parent in omath.iter():
                                if last_d in list(parent):
                                    parent.remove(last_d)
                                    print(f"  移除了公式内部的旧编号: ({texts[0]})")
                                    break
                    
                    wrap_in_eqarr_and_add_number(omath, effective_chapter, equation_number)
                    
                    # 设置 m:jc="left"，配合 eqArr + maxDist=1 + # 机制
                    m_ns = NAMESPACES['m']
                    oMathParaPr = math_para.find(f'{{{m_ns}}}oMathParaPr')
                    if oMathParaPr is None:
                        oMathParaPr = ET.SubElement(math_para, f'{{{m_ns}}}oMathParaPr')
                        math_para.insert(0, oMathParaPr)
                    jc = oMathParaPr.find(f'{{{m_ns}}}jc')
                    if jc is None:
                        jc = ET.SubElement(oMathParaPr, f'{{{m_ns}}}jc')
                    jc.set(f'{{{m_ns}}}val', 'left')
                    
                    modified_count += 1
                    
    print(f"成功修改 {modified_count} 个公式段落")
    return equation_map
