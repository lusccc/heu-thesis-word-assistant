"""
参考文献样式处理模块
功能：
1. 将参考文献列表的段落样式改为模板中的"参考文献"样式
2. 修复英文文献中的中文术语（等->et al., 卷->Vol等）
3. 修复中文文献中的英文术语（et al.->等, Vol->卷等）
   pandoc citeproc 不支持按 bib 条目 language 字段切换 locale，
   中文文献默认输出 et al.，需在后处理中修正
"""
import xml.etree.ElementTree as ET
import re
from .utils import NAMESPACES
from .config import BIBLIOGRAPHY_TITLE, STYLES

# 科研成果章节配置
PHD_OUTCOMES_TITLE = "读博士学位期间发表的论文和取得的科研成果"
PHD_OUTCOMES_SUBTITLE_PREFIX = "攻读"


def _is_english_reference(text):
    """判断是否为英文文献（基于作者名是否为英文大写字母）"""
    # 提取编号后的作者部分（如 "[1]AUTHOR NAME, ..."）
    match = re.match(r'^\[\d+\]\s*([A-Z][A-Z\s,]+)', text)
    if match:
        return True
    # 也检查是否包含大量英文字符
    english_chars = sum(1 for c in text if c.isascii() and c.isalpha())
    chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
    return english_chars > chinese_chars * 2


def _collect_non_hyperlink_text_groups(node, w_ns, groups, current_group):
    hyperlink_tag = f'{{{w_ns}}}hyperlink'
    text_tag = f'{{{w_ns}}}t'

    if node.tag == hyperlink_tag:
        if current_group:
            groups.append(current_group.copy())
            current_group.clear()
        return

    if node.tag == text_tag:
        current_group.append(node)

    for child in list(node):
        _collect_non_hyperlink_text_groups(child, w_ns, groups, current_group)


def _rewrite_text_group(t_elems, replacements):
    if not t_elems:
        return False

    full_text = ''.join(t.text or '' for t in t_elems)
    new_text = full_text
    for cn_term, en_term in replacements.items():
        if cn_term in new_text:
            new_text = new_text.replace(cn_term, en_term)

    if new_text == full_text:
        return False

    t_elems[0].text = new_text
    for t in t_elems[1:]:
        t.text = ''

    return True


def _fix_chinese_reference_terms(para, w_ns):
    """修复中文文献中的英文术语（pandoc citeproc 不按 language 切换 locale 的补救）"""
    replacements = {
        ', et al.': ', 等.',
        ', et al.,': ', 等,',
        'et al.': '等.',
        ': Vol ': ': 卷 ',
        ', Vol ': ', 卷 ',
        'Vol ': '卷 ',
    }

    groups = []
    current_group = []
    _collect_non_hyperlink_text_groups(para, w_ns, groups, current_group)
    if current_group:
        groups.append(current_group.copy())

    changed = False
    for t_elems in groups:
        if _rewrite_text_group(t_elems, replacements):
            changed = True

    return changed


def _fix_english_reference_terms(para, w_ns):
    """修复英文文献中的中文术语（在拼接完整文本上操作以处理跨run拆分）"""
    replacements = {
        ', 等.': ', et al.',
        ', 等,': ', et al.,',
        '等.': 'et al.',
        ': 卷 ': ': Vol ',
        ', 卷 ': ', Vol ',
        '卷 ': 'Vol ',
    }

    groups = []
    current_group = []
    _collect_non_hyperlink_text_groups(para, w_ns, groups, current_group)
    if current_group:
        groups.append(current_group.copy())

    changed = False
    for t_elems in groups:
        if _rewrite_text_group(t_elems, replacements):
            changed = True

    return changed


def apply_bibliography_style(root):
    """将参考文献列表的样式改为"参考文献"（af7），并修复英文文献术语"""
    print("正在处理参考文献样式...")
    
    w_ns = NAMESPACES['w']
    
    # 查找"参考文献"标题之后的所有段落
    in_bibliography = False
    style_count = 0
    term_fix_count = 0
    
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return
    
    for para in body.findall(f'{{{w_ns}}}p'):
        pPr = para.find(f'{{{w_ns}}}pPr')
        
        # 检查是否是"参考文献"标题
        if pPr is not None:
            pStyle = pPr.find(f'{{{w_ns}}}pStyle')
            if pStyle is not None:
                style_val = pStyle.get(f'{{{w_ns}}}val', '')
                if style_val in (STYLES['heading1'], STYLES['heading1_alt']):
                    # 提取标题文本
                    title_text = ''
                    for t in para.findall(f'.//{{{w_ns}}}t'):
                        if t.text:
                            title_text += t.text
                    
                    if BIBLIOGRAPHY_TITLE in title_text:
                        in_bibliography = True
                        continue
                    elif in_bibliography:
                        # 遇到下一个一级标题，结束参考文献区域
                        break
        
        # 如果在参考文献区域内，修改段落样式
        if in_bibliography:
            # 提取文本内容
            text_content = ''
            for t in para.findall(f'.//{{{w_ns}}}t'):
                if t.text:
                    text_content += t.text
            
            # 跳过空段落
            if not text_content.strip():
                continue
            
            # 确保有pPr
            if pPr is None:
                pPr = ET.Element(f'{{{w_ns}}}pPr')
                para.insert(0, pPr)
            
            # 设置样式
            pStyle = pPr.find(f'{{{w_ns}}}pStyle')
            if pStyle is None:
                pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
            
            # 设置为"参考文献"样式
            pStyle.set(f'{{{w_ns}}}val', STYLES['bibliography'])
            style_count += 1
            
            # 如果是英文文献，修复中文术语；如果是中文文献，修复英文术语
            if _is_english_reference(text_content):
                if _fix_english_reference_terms(para, w_ns):
                    term_fix_count += 1
            else:
                if _fix_chinese_reference_terms(para, w_ns):
                    term_fix_count += 1
    
    print(f"  已将 {style_count} 个参考文献条目设置为'参考文献'样式")
    print(f"  已修复 {term_fix_count} 个英文文献的中文术语")


def _set_heiti_size4(para, pPr, w_ns):
    """设置段落为黑体四号（14pt = 28半磅）"""
    if pPr is None:
        pPr = ET.Element(f'{{{w_ns}}}pPr')
        para.insert(0, pPr)

    # 设置段落默认字体属性
    rPr = pPr.find(f'{{{w_ns}}}rPr')
    if rPr is None:
        rPr = ET.SubElement(pPr, f'{{{w_ns}}}rPr')

    rFonts = rPr.find(f'{{{w_ns}}}rFonts')
    if rFonts is None:
        rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
    rFonts.set(f'{{{w_ns}}}eastAsia', '黑体')
    rFonts.set(f'{{{w_ns}}}ascii', '黑体')
    rFonts.set(f'{{{w_ns}}}hAnsi', '黑体')

    sz = rPr.find(f'{{{w_ns}}}sz')
    if sz is None:
        sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
    sz.set(f'{{{w_ns}}}val', '28')

    szCs = rPr.find(f'{{{w_ns}}}szCs')
    if szCs is None:
        szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
    szCs.set(f'{{{w_ns}}}val', '28')

    # 同时设置每个run的字体属性
    for run in para.findall(f'{{{w_ns}}}r'):
        run_rPr = run.find(f'{{{w_ns}}}rPr')
        if run_rPr is None:
            run_rPr = ET.Element(f'{{{w_ns}}}rPr')
            run.insert(0, run_rPr)

        run_rFonts = run_rPr.find(f'{{{w_ns}}}rFonts')
        if run_rFonts is None:
            run_rFonts = ET.SubElement(run_rPr, f'{{{w_ns}}}rFonts')
        run_rFonts.set(f'{{{w_ns}}}eastAsia', '黑体')
        run_rFonts.set(f'{{{w_ns}}}ascii', '黑体')
        run_rFonts.set(f'{{{w_ns}}}hAnsi', '黑体')

        run_sz = run_rPr.find(f'{{{w_ns}}}sz')
        if run_sz is None:
            run_sz = ET.SubElement(run_rPr, f'{{{w_ns}}}sz')
        run_sz.set(f'{{{w_ns}}}val', '28')

        run_szCs = run_rPr.find(f'{{{w_ns}}}szCs')
        if run_szCs is None:
            run_szCs = ET.SubElement(run_rPr, f'{{{w_ns}}}szCs')
        run_szCs.set(f'{{{w_ns}}}val', '28')


def apply_phd_outcomes_style(root):
    """处理"读博士学位期间发表的论文和取得的科研成果"章节的格式
    
    1. "攻读..."小标题设置为黑体四号
    2. [1][2]...条目应用"参考文献"样式，保留已有加粗格式
    3. 删除该章节中由空行产生的空段落
    """
    print("正在处理科研成果章节样式...")

    w_ns = NAMESPACES['w']
    in_section = False
    subtitle_count = 0
    item_count = 0
    empty_paras_to_remove = []

    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return

    for para in body.findall(f'{{{w_ns}}}p'):
        pPr = para.find(f'{{{w_ns}}}pPr')

        # 检查是否是一级标题
        if pPr is not None:
            pStyle = pPr.find(f'{{{w_ns}}}pStyle')
            if pStyle is not None:
                style_val = pStyle.get(f'{{{w_ns}}}val', '')
                if style_val in (STYLES['heading1'], STYLES['heading1_alt']):
                    title_text = ''
                    for t in para.findall(f'.//{{{w_ns}}}t'):
                        if t.text:
                            title_text += t.text

                    if PHD_OUTCOMES_TITLE in title_text:
                        in_section = True
                        continue
                    elif in_section:
                        break

        if not in_section:
            continue

        # 提取段落文本
        text_content = ''
        for t in para.findall(f'.//{{{w_ns}}}t'):
            if t.text:
                text_content += t.text
        text_content = text_content.strip()

        if not text_content:
            # 空段落：标记待删除（但保留包含分节符的段落）
            has_sectPr = (pPr is not None and pPr.find(f'{{{w_ns}}}sectPr') is not None)
            if not has_sectPr:
                empty_paras_to_remove.append(para)
            continue

        if text_content.startswith(PHD_OUTCOMES_SUBTITLE_PREFIX):
            # "攻读..."小标题：设置为黑体四号，删除首行缩进，段前1行
            _set_heiti_size4(para, pPr, w_ns)
            # 删除首行缩进
            if pPr is None:
                pPr = para.find(f'{{{w_ns}}}pPr')
            if pPr is not None:
                ind = pPr.find(f'{{{w_ns}}}ind')
                if ind is None:
                    ind = ET.SubElement(pPr, f'{{{w_ns}}}ind')
                ind.set(f'{{{w_ns}}}firstLineChars', '0')
                ind.set(f'{{{w_ns}}}firstLine', '0')
                # 清除可能存在的悬挂缩进
                if f'{{{w_ns}}}hanging' in ind.attrib:
                    del ind.attrib[f'{{{w_ns}}}hanging']
                # 段前1行
                spacing = pPr.find(f'{{{w_ns}}}spacing')
                if spacing is None:
                    spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
                spacing.set(f'{{{w_ns}}}beforeLines', '100')
            subtitle_count += 1
            print(f"  设置黑体四号(无缩进): {text_content[:30]}...")
        elif re.match(r'^\[\d+\]', text_content):
            # [1][2]...条目：设置悬挂缩进0.6厘米(340twips)，保留已有加粗
            if pPr is None:
                pPr = ET.Element(f'{{{w_ns}}}pPr')
                para.insert(0, pPr)
            ind = pPr.find(f'{{{w_ns}}}ind')
            if ind is None:
                ind = ET.SubElement(pPr, f'{{{w_ns}}}ind')
            ind.set(f'{{{w_ns}}}left', '340')
            ind.set(f'{{{w_ns}}}hanging', '340')
            ind.set(f'{{{w_ns}}}firstLineChars', '0')
            # 清除可能存在的首行缩进
            if f'{{{w_ns}}}firstLine' in ind.attrib:
                del ind.attrib[f'{{{w_ns}}}firstLine']
            item_count += 1

    # 删除空段落
    for para in empty_paras_to_remove:
        body.remove(para)

    print(f"  共设置 {subtitle_count} 个小标题为黑体四号")
    print(f"  共设置 {item_count} 个条目为悬挂缩进0.6cm")
    print(f"  共删除 {len(empty_paras_to_remove)} 个空段落")
