"""
图片样式处理模块
新方案：直接创建新段落替换原有的图片题注段落，避免Quarto生成的复杂段落结构问题
"""
import xml.etree.ElementTree as ET
import re
import os
from .utils import NAMESPACES, append_text_with_math
from .config import STYLES, DEFAULT_QMD_FILENAME, EXCLUDE_CHAPTER_TITLES


def unwrap_figure_layout_tables(root):
    """将图片从 Quarto 生成的布局表格中提取为独立段落

    Quarto/Pandoc 在生成 docx 时会将图片包裹在 1行1列 的表格中。
    本函数检测这类布局表格，将其内部的段落（图片+题注）提取到表格原位置，
    然后移除空的布局表格。
    """
    w_ns = NAMESPACES['w']
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return

    count = 0
    # 需要反复遍历，因为移除表格会改变子元素列表
    # 收集所有需要处理的表格及其位置
    to_process = []
    for idx, child in enumerate(list(body)):
        if child.tag != f'{{{w_ns}}}tbl':
            continue
        # 检查是否包含图片
        if child.find(f'.//{{{w_ns}}}drawing') is None:
            continue
        # 检查是否为简单的 1行1列 布局表格
        rows = child.findall(f'{{{w_ns}}}tr')
        if len(rows) != 1:
            continue
        cells = rows[0].findall(f'{{{w_ns}}}tc')
        if len(cells) != 1:
            continue
        to_process.append(child)

    for tbl in to_process:
        # 获取表格在 body 中的位置
        body_children = list(body)
        try:
            tbl_idx = body_children.index(tbl)
        except ValueError:
            continue

        # 提取单元格内的所有段落和书签元素（保持原始顺序）
        rows = tbl.findall(f'{{{w_ns}}}tr')
        cell = rows[0].find(f'{{{w_ns}}}tc')
        elements_to_extract = []
        for elem in cell:
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag in ('p', 'bookmarkStart', 'bookmarkEnd'):
                elements_to_extract.append(elem)

        if not elements_to_extract:
            continue

        # 在表格位置插入提取的元素（保持顺序）
        body.remove(tbl)
        for i, elem in enumerate(elements_to_extract):
            body.insert(tbl_idx + i, elem)

        count += 1

    if count > 0:
        print(f"  从 {count} 个布局表格中提取了图片为独立段落")


def load_en_captions_from_qmd(qmd_path=None):
    """从thesis.qmd读取英文图片标题"""
    if qmd_path is None:
        qmd_path = DEFAULT_QMD_FILENAME
    
    captions = {}
    if not os.path.exists(qmd_path):
        return captions
    
    with open(qmd_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 匹配 {#fig-xxx} 后面的 <!-- fig-cap-en: ... -->
    pattern = r'\{#(fig-[^\s}]+)[^}]*\}[^\n]*\n\s*<!-- fig-cap-en:\s*(.+?)\s*-->'
    for match in re.finditer(pattern, content):
        fig_id = match.group(1)
        en_caption = match.group(2)
        captions[fig_id] = en_caption
    
    return captions


def create_caption_paragraph(w_ns, text, style_id):
    """
    创建一个干净的题注段落
    显式设置段前段后间距为0，避免Word使用默认值
    支持 $...$ 中的LaTeX公式，转换为OMML数学元素
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
    
    jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
    jc.set(f'{{{w_ns}}}val', 'center')
    
    # 文本内容（支持 $...$ 公式）
    append_text_with_math(para, w_ns, text)
    
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


def apply_figure_style_to_image_paragraphs(root):
    """
    为包含图片的段落应用"图"样式
    查找所有包含 w:drawing 元素的段落，将其样式设置为模板中的"图"样式
    """
    w_ns = NAMESPACES['w']
    figure_style_id = STYLES.get('figure', 'af6')
    
    count = 0
    for para in root.iter(f'{{{w_ns}}}p'):
        # 检查段落是否包含图片（w:drawing 元素）
        has_drawing = para.find(f'.//{{{w_ns}}}drawing') is not None
        if not has_drawing:
            continue
        
        # 获取或创建段落属性
        pPr = para.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            pPr = ET.Element(f'{{{w_ns}}}pPr')
            para.insert(0, pPr)
        
        # 设置或更新段落样式
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
        pStyle.set(f'{{{w_ns}}}val', figure_style_id)
        
        count += 1
    
    print(f"  为 {count} 个图片段落应用了'图'样式")
    return count


def process_figures(root, qmd_path=None):
    """处理图片样式和题注"""
    print("正在处理图片样式...")
    
    # 首先将图片从布局表格中提取为独立段落
    unwrap_figure_layout_tables(root)
    
    # 为所有包含图片的段落应用"图"样式
    apply_figure_style_to_image_paragraphs(root)
    
    # 从qmd文件读取英文标题
    en_captions = load_en_captions_from_qmd(qmd_path)
    print(f"  从qmd文件读取到 {len(en_captions)} 个英文标题")
    
    w_ns = NAMESPACES['w']
    
    # 查找所有fig-书签，建立顺序映射（从root全局搜索，因书签可能不在w:p内）
    fig_bookmarks = []
    for bm in root.iter(f'{{{w_ns}}}bookmarkStart'):
        name = bm.get(f'{{{w_ns}}}name', '')
        if name.startswith('fig-'):
            fig_bookmarks.append(name)
    
    # 动态检测章节并收集图片题注
    # 遍历文档段落，跟踪当前章节号，为每个图片分配正确的章节
    current_chapter = 0
    caption_paras_info = []  # (para, fig_seq, cn_caption, chapter)
    fig_seq = 0
    
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
        
        # 匹配图片题注
        text = get_paragraph_text(para, w_ns)
        match = re.match(r'^Figure\s*(\d+(?:\.\d+)?):\s*(.*)$', text)
        if match:
            fig_seq += 1
            cn_caption = match.group(2)
            caption_paras_info.append((para, fig_seq, cn_caption, current_chapter))
    
    print(f"  找到 {len(caption_paras_info)} 个图片")
    
    # 处理每个图片题注
    figure_map = {}
    chapter_fig_count = {}
    
    for para, fig_num_orig, cn_caption, chapter in caption_paras_info:
        if fig_num_orig > len(fig_bookmarks):
            continue
        
        fig_id = fig_bookmarks[fig_num_orig - 1]
        
        chapter_fig_count[chapter] = chapter_fig_count.get(chapter, 0) + 1
        fig_num = chapter_fig_count[chapter]
        figure_map[fig_id] = (chapter, fig_num)
        
        # 获取英文标题
        en_caption = en_captions.get(fig_id, cn_caption)
        
        new_prefix = f"图{chapter}.{fig_num} "
        
        # === 替换 "Figure X:" 前缀（处理可能跨多个 run 的情况） ===
        # 收集 pPr 之后的所有 w:r 及其文本，定位前缀边界
        prefix_pattern = re.compile(r'^Figure\s*\d+(?:\.\d+)?:\s*')
        child_runs = []  # (element, t_elem, run_text, start_pos)
        accumulated = ''
        for child in list(para):
            if child.tag == f'{{{w_ns}}}r':
                t = child.find(f'{{{w_ns}}}t')
                run_text = (t.text or '') if t is not None else ''
                child_runs.append((child, t, run_text, len(accumulated)))
                accumulated += run_text
        
        prefix_match = prefix_pattern.match(accumulated)
        if prefix_match:
            prefix_end = prefix_match.end()
            # 移除/截断被前缀覆盖的 run
            for run_elem, t_elem, run_text, start_pos in child_runs:
                end_pos = start_pos + len(run_text)
                if end_pos <= prefix_end:
                    # 整个 run 在前缀范围内，移除
                    para.remove(run_elem)
                elif start_pos < prefix_end:
                    # run 部分在前缀范围内，截断保留后面部分
                    keep_from = prefix_end - start_pos
                    if t_elem is not None:
                        t_elem.text = run_text[keep_from:]
                    break
                else:
                    break
            # 在 pPr 后插入新前缀 run
            insert_pos = 0
            for idx_c, c in enumerate(para):
                if c.tag == f'{{{w_ns}}}pPr':
                    insert_pos = idx_c + 1
                    break
            new_run = ET.Element(f'{{{w_ns}}}r')
            new_t = ET.SubElement(new_run, f'{{{w_ns}}}t')
            new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            new_t.text = new_prefix
            para.insert(insert_pos, new_run)
        
        # === 完全替换 pPr，确保样式、居中、间距正确 ===
        old_pPr = para.find(f'{{{w_ns}}}pPr')
        if old_pPr is not None:
            para.remove(old_pPr)
        
        pPr = ET.Element(f'{{{w_ns}}}pPr')
        para.insert(0, pPr)
        
        pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
        pStyle.set(f'{{{w_ns}}}val', STYLES['figure_caption'])
        
        spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
        spacing.set(f'{{{w_ns}}}after', '0')
        spacing.set(f'{{{w_ns}}}before', '0')
        
        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
        jc.set(f'{{{w_ns}}}val', 'center')
        
        # 在原段落后面插入英文题注段落
        en_text = f"Fig.{chapter}.{fig_num} {en_caption}"
        new_en_para = create_caption_paragraph(w_ns, en_text, STYLES['figure_caption'])
        
        parent = find_parent(root, para)
        if parent is not None:
            idx = list(parent).index(para)
            parent.insert(idx + 1, new_en_para)
        
        print(f"    图{chapter}.{fig_num}: {cn_caption}")
    
    # 更新交叉引用
    update_figure_references(root, figure_map)
    
    return figure_map


def update_figure_references(root, figure_map):
    """更新图片交叉引用"""
    w_ns = NAMESPACES['w']
    refs_count = 0
    
    for para in root.iter(f'{{{w_ns}}}p'):
        children = list(para)
        to_remove = []
        
        for i, child in enumerate(children):
            if child.tag == f'{{{w_ns}}}hyperlink':
                anchor = child.get(f'{{{w_ns}}}anchor')
                if anchor and anchor in figure_map:
                    ch, fn = figure_map[anchor]
                    new_text = f"图{ch}.{fn}"
                    
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
                    
                    # 移除前面的"图"和空格
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
                                elif text.endswith('图'):
                                    t.text = text[:-1]
                                    break
                                elif text.endswith('图 '):
                                    t.text = text[:-2]
                                    break
                                elif text.strip() == '图':
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
    
    print(f"  更新了 {refs_count} 个图片交叉引用")
