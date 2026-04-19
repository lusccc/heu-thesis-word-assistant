"""
摘要样式处理模块
功能：
1. 将摘要和Abstract移到目录之前
2. Abstract标题加粗
"""
import xml.etree.ElementTree as ET
from .utils import NAMESPACES
from .config import STYLES


def move_abstract_before_toc(root):
    """将摘要和Abstract移到目录(sdt)之前"""
    print("正在调整摘要位置...")
    
    w_ns = NAMESPACES['w']
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        print("  未找到document body")
        return
    
    children = list(body)
    
    # 1. 找到sdt（目录）
    sdt_index = None
    for i, child in enumerate(children):
        if child.tag.split('}')[-1] == 'sdt':
            sdt_index = i
            break
    
    if sdt_index is None:
        print("  未找到目录sdt，跳过")
        return
    
    # 2. 找到sdt后面的第一个包含sectPr的段落（目录分节符）
    toc_sect_index = None
    for i in range(sdt_index + 1, len(children)):
        child = children[i]
        if child.tag.split('}')[-1] != 'p':
            continue
        pPr = child.find(f'{{{w_ns}}}pPr')
        if pPr is not None and pPr.find(f'{{{w_ns}}}sectPr') is not None:
            toc_sect_index = i
            break
    
    if toc_sect_index is None:
        print("  未找到目录分节符，跳过")
        return
    
    # 3. 从目录分节符后面开始，找到第一个非摘要/非Abstract的一级标题
    first_content_index = None
    for i in range(toc_sect_index + 1, len(children)):
        child = children[i]
        if child.tag.split('}')[-1] != 'p':
            continue
        pPr = child.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue
        if pStyle.get(f'{{{w_ns}}}val', '') not in (STYLES['heading1'], STYLES['heading1_alt']):
            continue
        text = ''
        for t in child.findall(f'.//{{{w_ns}}}t'):
            if t.text:
                text += t.text
        text_nospace = text.strip().replace(' ', '').replace('\u3000', '')
        if '摘要' in text_nospace or text_nospace.lower() == 'abstract' or '创新成果' in text_nospace:
            continue
        first_content_index = i
        break
    
    if first_content_index is None:
        print("  未找到正文章节标题，跳过")
        return
    
    # 4. 确定要移动的范围
    start_index = toc_sect_index + 1
    end_index = first_content_index - 1
    
    # 排除末尾属于下一章节的bookmarkStart
    while end_index >= start_index:
        child = children[end_index]
        if child.tag.split('}')[-1] == 'bookmarkStart':
            end_index -= 1
        else:
            break
    
    if start_index > end_index:
        print("  摘要已在目录之前，无需移动")
        return
    
    # 5. 收集、移除、重新插入
    elements_to_move = [children[i] for i in range(start_index, end_index + 1)]
    
    for elem in elements_to_move:
        body.remove(elem)
    
    # 重新找到sdt位置
    new_children = list(body)
    new_sdt_index = None
    for i, child in enumerate(new_children):
        if child.tag.split('}')[-1] == 'sdt':
            new_sdt_index = i
            break
    
    for j, elem in enumerate(elements_to_move):
        body.insert(new_sdt_index + j, elem)
    
    print(f"  已将摘要和Abstract移到目录之前（移动了{len(elements_to_move)}个元素）")


def exclude_abstract_from_toc(root):
    """通过TOC域代码的\\b开关+书签范围，排除摘要和Abstract出现在目录中
    
    方案：
    1. 修改sdt中的TOC域代码，添加 \\b _TocRange 开关
    2. 在目录后的分节符段落之后添加bookmarkStart(_TocRange)
    3. 在body末尾的sectPr之前添加bookmarkEnd(_TocRange)
    这样TOC只收集_TocRange书签范围内的标题，摘要和Abstract在范围之外。
    """
    print("正在排除摘要/Abstract出现在TOC中...")
    
    w_ns = NAMESPACES['w']
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return
    
    children = list(body)
    BOOKMARK_NAME = '_TocRange'
    
    # 1. 找到sdt并修改TOC域代码
    sdt = None
    sdt_index = None
    for i, child in enumerate(children):
        if child.tag.split('}')[-1] == 'sdt':
            sdt = child
            sdt_index = i
            break
    
    if sdt is None:
        print("  未找到目录sdt，跳过")
        return
    
    # 修改instrText
    modified = False
    for instr in sdt.findall(f'.//{{{w_ns}}}instrText'):
        if instr.text and 'TOC' in instr.text and '\\b' not in instr.text:
            instr.text = instr.text.replace('TOC ', f'TOC \\b {BOOKMARK_NAME} ', 1)
            modified = True
            print(f"  修改TOC域代码: {instr.text.strip()}")
    
    if not modified:
        print("  TOC域代码已包含\\b开关或未找到，跳过")
        return
    
    # 2. 找到目录后的分节符段落，在其后插入bookmarkStart
    # 分配一个不冲突的bookmark ID
    max_bm_id = 0
    for bm in root.findall(f'.//{{{w_ns}}}bookmarkStart'):
        try:
            bm_id = int(bm.get(f'{{{w_ns}}}id', '0'))
            max_bm_id = max(max_bm_id, bm_id)
        except ValueError:
            pass
    new_bm_id = str(max_bm_id + 1)
    
    # 找到第一个正文章节标题（不在EXCLUDE_CHAPTER_TITLES中的一级标题）
    # bookmarkStart插入到该标题之前，这样目录和创新成果自评表等都在范围之外
    from .config import EXCLUDE_CHAPTER_TITLES
    
    first_content_heading_index = None
    for i in range(sdt_index + 1, len(children)):
        child = children[i]
        if child.tag.split('}')[-1] != 'p':
            continue
        pPr = child.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue
        if pStyle.get(f'{{{w_ns}}}val', '') not in (STYLES['heading1'], STYLES['heading1_alt']):
            continue
        # 是一级标题，检查是否在排除列表中
        h_text = ''
        for t in child.findall(f'.//{{{w_ns}}}t'):
            if t.text:
                h_text += t.text
        h_text_nospace = h_text.strip().replace(' ', '').replace('\u3000', '')
        is_excluded = any(exc in h_text_nospace for exc in EXCLUDE_CHAPTER_TITLES)
        if not is_excluded:
            first_content_heading_index = i
            print(f"  第一个正文章节: {h_text.strip()[:30]}")
            break
    
    if first_content_heading_index is None:
        print("  未找到正文章节标题，跳过")
        return
    
    # 在第一个正文章节标题之前插入bookmarkStart
    # 需要找到该标题前面的分节符段落（如果有的话），bookmarkStart插在分节符之后
    # 或者直接插在标题之前
    insert_index = first_content_heading_index
    # 检查标题前面是否有bookmarkStart（如sec-intro），如果有则插在bookmarkStart之前
    while insert_index > 0 and children[insert_index - 1].tag.split('}')[-1] == 'bookmarkStart':
        insert_index -= 1
    
    bm_start = ET.Element(f'{{{w_ns}}}bookmarkStart')
    bm_start.set(f'{{{w_ns}}}id', new_bm_id)
    bm_start.set(f'{{{w_ns}}}name', BOOKMARK_NAME)
    body.insert(insert_index, bm_start)
    
    # 3. 在body末尾的sectPr之前插入bookmarkEnd
    # 重新获取children（因为插入了元素）
    children = list(body)
    last_sectPr_index = None
    for i in range(len(children) - 1, -1, -1):
        if children[i].tag.split('}')[-1] == 'sectPr':
            last_sectPr_index = i
            break
    
    if last_sectPr_index is not None:
        bm_end = ET.Element(f'{{{w_ns}}}bookmarkEnd')
        bm_end.set(f'{{{w_ns}}}id', new_bm_id)
        body.insert(last_sectPr_index, bm_end)
        print(f"  已添加书签 {BOOKMARK_NAME} (id={new_bm_id})，TOC将只收集书签范围内的标题")


def bold_abstract_title(root):
    """给Abstract标题加粗"""
    print("正在处理Abstract标题加粗...")
    
    w_ns = NAMESPACES['w']
    
    for para in root.findall(f'.//{{{w_ns}}}p'):
        pPr = para.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue
        if pStyle.get(f'{{{w_ns}}}val', '') not in (STYLES['heading1'], STYLES['heading1_alt']):
            continue
        
        text = ''
        for t in para.findall(f'.//{{{w_ns}}}t'):
            if t.text:
                text += t.text
        
        if text.strip().lower() == 'abstract':
            for run in para.findall(f'{{{w_ns}}}r'):
                rPr = run.find(f'{{{w_ns}}}rPr')
                if rPr is None:
                    rPr = ET.Element(f'{{{w_ns}}}rPr')
                    run.insert(0, rPr)
                b = rPr.find(f'{{{w_ns}}}b')
                if b is None:
                    ET.SubElement(rPr, f'{{{w_ns}}}b')
            print("  已给Abstract标题加粗")
            break


def set_keywords_heiti(root):
    """将中文摘要中"关键词："的字体设为黑体（SimHei）"""
    print("正在处理关键词黑体字体...")
    
    w_ns = NAMESPACES['w']
    
    for para in root.findall(f'.//{{{w_ns}}}p'):
        for run in para.findall(f'{{{w_ns}}}r'):
            run_text = ''
            for t in run.findall(f'{{{w_ns}}}t'):
                if t.text:
                    run_text += t.text
            
            if '关键词' in run_text:
                rPr = run.find(f'{{{w_ns}}}rPr')
                if rPr is None:
                    rPr = ET.Element(f'{{{w_ns}}}rPr')
                    run.insert(0, rPr)
                rFonts = rPr.find(f'{{{w_ns}}}rFonts')
                if rFonts is None:
                    rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
                rFonts.set(f'{{{w_ns}}}ascii', 'SimHei')
                rFonts.set(f'{{{w_ns}}}eastAsia', 'SimHei')
                rFonts.set(f'{{{w_ns}}}hAnsi', 'SimHei')
                print(f"  已设置黑体: {run_text.strip()}")
                return
    
    print("  未找到关键词文本")


def apply_innovation_table_borders(root):
    """给"博士学位论文创新成果自评表"设置全边框
    
    识别方法：找到第一行包含"序号"和"创新"的表格。
    全边框：每个单元格的top/bottom/left/right都设为single线。
    """
    print("正在处理创新成果自评表边框...")
    
    w_ns = NAMESPACES['w']
    
    for tbl in root.findall(f'.//{{{w_ns}}}tbl'):
        rows = list(tbl.findall(f'{{{w_ns}}}tr'))
        if not rows:
            continue
        
        # 检查第一行文本
        first_row_text = ''
        for t in rows[0].findall(f'.//{{{w_ns}}}t'):
            if t.text:
                first_row_text += t.text
        
        if '序号' not in first_row_text or '创新' not in first_row_text:
            continue
        
        # 找到创新成果表，设置全边框
        border_attrs = {
            f'{{{w_ns}}}val': 'single',
            f'{{{w_ns}}}sz': '4',
            f'{{{w_ns}}}space': '0',
            f'{{{w_ns}}}color': '000000',
        }
        
        for tr in rows:
            for tc in tr.findall(f'{{{w_ns}}}tc'):
                tcPr = tc.find(f'{{{w_ns}}}tcPr')
                if tcPr is None:
                    tcPr = ET.Element(f'{{{w_ns}}}tcPr')
                    tc.insert(0, tcPr)
                
                # 移除旧边框
                old_b = tcPr.find(f'{{{w_ns}}}tcBorders')
                if old_b is not None:
                    tcPr.remove(old_b)
                
                tcBorders = ET.SubElement(tcPr, f'{{{w_ns}}}tcBorders')
                for direction in ('top', 'left', 'bottom', 'right'):
                    elem = ET.SubElement(tcBorders, f'{{{w_ns}}}{direction}')
                    for attr, value in border_attrs.items():
                        elem.set(attr, value)
        
        # 同时设置表格级别边框
        tblPr = tbl.find(f'{{{w_ns}}}tblPr')
        if tblPr is not None:
            old_tb = tblPr.find(f'{{{w_ns}}}tblBorders')
            if old_tb is not None:
                tblPr.remove(old_tb)
            tblBorders = ET.SubElement(tblPr, f'{{{w_ns}}}tblBorders')
            for direction in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                elem = ET.SubElement(tblBorders, f'{{{w_ns}}}{direction}')
                for attr, value in border_attrs.items():
                    elem.set(attr, value)
        
        # 缩小第一列（序号列）宽度
        COL1_WIDTH = '400'  # 约0.7cm
        tblGrid = tbl.find(f'{{{w_ns}}}tblGrid')
        if tblGrid is not None:
            gridCols = list(tblGrid.findall(f'{{{w_ns}}}gridCol'))
            if len(gridCols) >= 2:
                gridCols[0].set(f'{{{w_ns}}}w', COL1_WIDTH)
        
        # 22磅行距 = 440 twips (1磅=20twips)
        LINE_SPACING = '440'
        # 首行缩进2字符 ≈ 420 twips (小四号字体约210twips/字符)
        FIRST_LINE_INDENT = '420'
        
        for row_idx, tr in enumerate(rows):
            # 允许跨页断行 & 不重复表头
            trPr = tr.find(f'{{{w_ns}}}trPr')
            if trPr is None:
                trPr = ET.Element(f'{{{w_ns}}}trPr')
                tr.insert(0, trPr)
            for tag in ('cantSplit', 'tblHeader'):
                old = trPr.find(f'{{{w_ns}}}{tag}')
                if old is not None:
                    trPr.remove(old)
            
            tcs = list(tr.findall(f'{{{w_ns}}}tc'))
            
            for col_idx, tc in enumerate(tcs):
                # 设置第一列tcW
                if col_idx == 0:
                    tcPr = tc.find(f'{{{w_ns}}}tcPr')
                    if tcPr is not None:
                        tcW = tcPr.find(f'{{{w_ns}}}tcW')
                        if tcW is None:
                            tcW = ET.SubElement(tcPr, f'{{{w_ns}}}tcW')
                        tcW.set(f'{{{w_ns}}}w', COL1_WIDTH)
                        tcW.set(f'{{{w_ns}}}type', 'dxa')
                
                # 设置段落格式
                for para in tc.findall(f'{{{w_ns}}}p'):
                    pPr = para.find(f'{{{w_ns}}}pPr')
                    if pPr is None:
                        pPr = ET.Element(f'{{{w_ns}}}pPr')
                        para.insert(0, pPr)
                    
                    # 行间距22磅（固定值）
                    spacing = pPr.find(f'{{{w_ns}}}spacing')
                    if spacing is None:
                        spacing = ET.SubElement(pPr, f'{{{w_ns}}}spacing')
                    spacing.set(f'{{{w_ns}}}line', LINE_SPACING)
                    spacing.set(f'{{{w_ns}}}lineRule', 'exact')
                    
                    # 对齐和缩进
                    # 移除已有的jc和ind
                    for old_jc in list(pPr.findall(f'{{{w_ns}}}jc')):
                        pPr.remove(old_jc)
                    for old_ind in list(pPr.findall(f'{{{w_ns}}}ind')):
                        pPr.remove(old_ind)
                    
                    if row_idx == 0:
                        # 表头行：居中
                        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
                        jc.set(f'{{{w_ns}}}val', 'center')
                    elif col_idx == 0:
                        # 内容行第一列（序号）：居中
                        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
                        jc.set(f'{{{w_ns}}}val', 'center')
                    else:
                        # 内容行第二列：居左，首行缩进2字符
                        jc = ET.SubElement(pPr, f'{{{w_ns}}}jc')
                        jc.set(f'{{{w_ns}}}val', 'left')
                        ind = ET.SubElement(pPr, f'{{{w_ns}}}ind')
                        ind.set(f'{{{w_ns}}}firstLine', FIRST_LINE_INDENT)
        
        print(f"  已设置全边框（{len(rows)}行），序号列宽{COL1_WIDTH}twips，22磅行距，首行缩进")
        return
    
    print("  未找到创新成果自评表")
