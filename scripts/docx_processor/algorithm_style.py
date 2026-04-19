"""
算法样式处理模块
处理算法块：
1. 自动按章节编号（类似图表）
2. 中英文双标题（参考图表标题样式，不加粗）
3. 三线表形式段落边框（"输入"上方顶线、"输出"下方中间线、最后一步下方底线）
4. 支持交叉引用（更新全文中"算法X.Y"的编号）
"""
import xml.etree.ElementTree as ET
import re
import os
from .utils import NAMESPACES, append_text_with_math
from .config import STYLES, DEFAULT_QMD_FILENAME, EXCLUDE_CHAPTER_TITLES


def _read_algo_info_from_qmd(qmd_path):
    """从QMD文件按出现顺序读取算法标题信息
    
    格式：
        <!-- algo-title: 中文标题 -->
        <!-- algo-cap-en: 英文标题 -->
    
    Returns:
        list: [(cn_title, en_title)] 按出现顺序
    """
    if not os.path.exists(qmd_path):
        return []
    
    with open(qmd_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    cn_titles = [m.group(1).strip() for m in re.finditer(r'<!--\s*algo-title:\s*(.+?)\s*-->', content)]
    en_titles = [m.group(1).strip() for m in re.finditer(r'<!--\s*algo-cap-en:\s*(.+?)\s*-->', content)]
    
    # 去掉英文标题中可能残留的编号前缀
    cleaned_en = []
    for t in en_titles:
        m = re.match(r'Algorithm\s+\d+\.\d+\s+(.*)', t)
        cleaned_en.append(m.group(1) if m else t)
    
    # 按顺序配对
    result = []
    for i in range(max(len(cn_titles), len(cleaned_en))):
        cn = cn_titles[i] if i < len(cn_titles) else ''
        en = cleaned_en[i] if i < len(cleaned_en) else ''
        result.append((cn, en))
    
    return result


def _add_paragraph_border(para, w_ns, position, sz='12', color='000000'):
    """给段落添加边框"""
    pPr = para.find(f'{{{w_ns}}}pPr')
    if pPr is None:
        pPr = ET.SubElement(para, f'{{{w_ns}}}pPr')
        para.remove(pPr)
        para.insert(0, pPr)
    
    pBdr = pPr.find(f'{{{w_ns}}}pBdr')
    if pBdr is None:
        pBdr = ET.SubElement(pPr, f'{{{w_ns}}}pBdr')
    
    border = ET.SubElement(pBdr, f'{{{w_ns}}}{position}')
    border.set(f'{{{w_ns}}}val', 'single')
    border.set(f'{{{w_ns}}}sz', sz)
    border.set(f'{{{w_ns}}}space', '1')
    border.set(f'{{{w_ns}}}color', color)


def _create_caption_paragraph(w_ns, text, style_id):
    """创建算法标题段落（参考图表标题样式：居中、无间距、不加粗）"""
    para = ET.Element(f'{{{w_ns}}}p')
    
    pPr = ET.SubElement(para, f'{{{w_ns}}}pPr')
    pStyle = ET.SubElement(pPr, f'{{{w_ns}}}pStyle')
    pStyle.set(f'{{{w_ns}}}val', style_id)
    
    spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
    spacing.set(f'{{{w_ns}}}after', '0')
    spacing.set(f'{{{w_ns}}}before', '0')
    
    jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
    jc.set(f'{{{w_ns}}}val', 'center')
    
    append_text_with_math(para, w_ns, text)
    
    return para


def _get_paragraph_text(para, w_ns):
    """获取段落纯文本"""
    text = ''
    for t in para.iter(f'{{{w_ns}}}t'):
        if t.text:
            text += t.text
    return text


def _update_algorithm_references(root, algo_bookmark_map):
    """更新全文中 algo- hyperlink 的交叉引用文本
    
    与图表交叉引用机制一致：遍历所有 hyperlink，
    如果 anchor 是 algo- 书签，则更新其内部 run 的文本为"算法X.Y"。
    
    Args:
        root: 文档根元素
        algo_bookmark_map: {书签名: 新编号} 如 {'algo-imputation': '4.1'}
    """
    w_ns = NAMESPACES['w']
    if not algo_bookmark_map:
        return
    
    updated = 0
    for para in root.iter(f'{{{w_ns}}}p'):
        for child in list(para):
            if child.tag != f'{{{w_ns}}}hyperlink':
                continue
            anchor = child.get(f'{{{w_ns}}}anchor', '')
            if not anchor.startswith('algo-'):
                continue
            if anchor not in algo_bookmark_map:
                continue
            
            new_id = algo_bookmark_map[anchor]
            new_text = f'算法{new_id}'
            
            # 更新 hyperlink 内所有 run 的文本：
            # 第一个 run 设为新文本，其余 run 的文本清空
            runs = child.findall(f'{{{w_ns}}}r')
            for idx_r, run in enumerate(runs):
                t = run.find(f'{{{w_ns}}}t')
                if t is not None:
                    if idx_r == 0:
                        t.text = new_text
                    else:
                        t.text = ''
            updated += 1
    
    if updated > 0:
        print(f"  更新了 {updated} 个算法交叉引用")


def process_algorithms(root, qmd_path):
    """处理算法块样式
    
    从 QMD 的 <!-- algo-title: ... --> 和 <!-- algo-cap-en: ... --> 注释读取标题。
    在 docx 中通过"输入："段落定位算法块，自动编号并插入中英文标题。
    
    三线表边框：
    - 顶线："输入："段落上边框
    - 中间线："输出："段落下边框
    - 底线：最后一步段落下边框
    """
    w_ns = NAMESPACES['w']
    caption_style = STYLES['figure_caption']
    
    print("正在处理算法样式...")
    
    # 从 QMD 读取算法标题信息
    algo_titles = _read_algo_info_from_qmd(qmd_path)
    if not algo_titles:
        print("  未在QMD中找到算法标题定义")
        return
    
    print(f"  从QMD读取到 {len(algo_titles)} 个算法标题")
    
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        print("  未找到body元素")
        return
    
    # === 遍历 body 直接子元素，检测章节并收集 algo- 书签和"输入："段落 ===
    current_chapter = 0
    input_paras = []  # [(input_para, chapter, bookmark_name)]
    pending_algo_bookmark = None  # 暂存待关联的 algo- 书签名
    
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        
        # 收集 algo- 书签（body 直接子元素，Pandoc 为 div ID 生成）
        if tag == 'bookmarkStart':
            name = child.get(f'{{{w_ns}}}name', '')
            if name.startswith('algo-'):
                pending_algo_bookmark = name
            continue
        
        if tag != 'p':
            continue
        
        # 检测章节标题
        pPr = child.find(f'{{{w_ns}}}pPr')
        if pPr is not None:
            pStyle = pPr.find(f'{{{w_ns}}}pStyle')
            if pStyle is not None:
                style_val = pStyle.get(f'{{{w_ns}}}val', '')
                if style_val in (STYLES['heading1'], STYLES.get('heading1_alt', 'Heading1')):
                    para_text = _get_paragraph_text(child, w_ns).strip()
                    para_text_nospace = para_text.replace(' ', '').replace('\u3000', '')
                    is_excluded = any(ex in para_text_nospace for ex in EXCLUDE_CHAPTER_TITLES)
                    if not is_excluded:
                        current_chapter += 1
        
        # 也检查段落内部的 algo- 书签
        for bm in child.findall(f'.//{{{w_ns}}}bookmarkStart'):
            name = bm.get(f'{{{w_ns}}}name', '')
            if name.startswith('algo-'):
                pending_algo_bookmark = name
        
        # 匹配"输入："段落（算法块的起始标志）
        text = _get_paragraph_text(child, w_ns).strip()
        if text.startswith('输入：') or text.startswith('输入:'):
            bm_name = pending_algo_bookmark if pending_algo_bookmark else ''
            input_paras.append((child, current_chapter, bm_name))
            pending_algo_bookmark = None
    
    # 只取与 QMD 中算法数量匹配的"输入："段落
    if len(input_paras) < len(algo_titles):
        print(f"  警告：docx中找到 {len(input_paras)} 个'输入：'段落，少于QMD中 {len(algo_titles)} 个算法")
    
    algo_count = min(len(input_paras), len(algo_titles))
    
    # === 处理每个算法 ===
    chapter_algo_count = {}
    algo_bookmark_map = {}  # 书签名 → 新编号
    
    for seq in range(algo_count):
        input_para, chapter, bm_name = input_paras[seq]
        cn_title, en_title = algo_titles[seq]
        
        chapter_algo_count[chapter] = chapter_algo_count.get(chapter, 0) + 1
        algo_num = chapter_algo_count[chapter]
        new_id = f"{chapter}.{algo_num}"
        
        # 记录书签到编号的映射（用于交叉引用）
        if bm_name:
            algo_bookmark_map[bm_name] = new_id
            print(f"  算法{new_id}: {cn_title[:30]}... (书签: {bm_name})")
        else:
            print(f"  算法{new_id}: {cn_title[:30]}...")
        
        # === 1. 在"输入："段落前插入中英文标题 ===
        cn_text = f"算法{new_id} {cn_title}"
        new_cn_para = _create_caption_paragraph(w_ns, cn_text, caption_style)
        
        en_text = f"Algorithm {new_id} {en_title}"
        new_en_para = _create_caption_paragraph(w_ns, en_text, caption_style)
        
        input_idx = list(body).index(input_para)
        body.insert(input_idx, new_cn_para)
        body.insert(input_idx + 1, new_en_para)
        
        # === 2. 向下扫描找到"输出："和最后一步 ===
        paragraphs = list(body)
        input_new_idx = paragraphs.index(input_para)
        
        output_para = None
        last_step_para = None
        
        for j in range(input_new_idx, min(input_new_idx + 60, len(paragraphs))):
            p = paragraphs[j]
            if p.tag != f'{{{w_ns}}}p':
                continue
            p_text = _get_paragraph_text(p, w_ns).strip()
            
            if p_text.startswith('输出：') or p_text.startswith('输出:'):
                output_para = p
            
            if re.match(r'^步骤\d+[：:]', p_text):
                last_step_para = p
            
            # 遇到标题样式说明算法块已结束
            pPr_j = p.find(f'{{{w_ns}}}pPr')
            if pPr_j is not None:
                pStyle_j = pPr_j.find(f'{{{w_ns}}}pStyle')
                if pStyle_j is not None:
                    style_j = pStyle_j.get(f'{{{w_ns}}}val', '')
                    if style_j and (style_j.startswith('Heading') or style_j in ('1', '2', '3', '4')):
                        break
            
            if re.match(r'^步骤\d+[：:]\s*结束', p_text):
                last_step_para = p
                break
        
        # === 3. 添加三线表边框 ===
        _add_paragraph_border(input_para, w_ns, 'top', sz='12')
        
        if output_para is not None:
            _add_paragraph_border(output_para, w_ns, 'bottom', sz='4')
        
        if last_step_para is not None:
            _add_paragraph_border(last_step_para, w_ns, 'bottom', sz='12')
    
    # === 更新全文交叉引用 ===
    if algo_bookmark_map:
        print(f"  算法书签映射: {algo_bookmark_map}")
        _update_algorithm_references(root, algo_bookmark_map)
    
    print(f"  共处理 {algo_count} 个算法块")
