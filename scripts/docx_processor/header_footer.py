"""
页眉页脚处理模块
功能：
1. 目录之后（正文章节）奇数页页眉：第X章 章节标题
2. 目录及之前（摘要/目录）奇数页页眉：学位论文题目
3. 偶数页页眉：学校名称+论文类型
4. 页脚：居中页码
5. 封面页（title page）无页眉，由外部docx自带设置
"""
import xml.etree.ElementTree as ET
from pathlib import Path
from .config import (
    UNIVERSITY_NAME, THESIS_TYPE, 
    EXCLUDE_CHAPTER_TITLES, STYLES
)
from .toc_style import TOC_TITLE

# 命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# 偶数页页眉文本
EVEN_HEADER_TEXT = f"{UNIVERSITY_NAME}{THESIS_TYPE}"


def create_header_xml(text):
    """创建页眉XML内容
    
    使用模板中的"ae"样式（页眉样式）
    """
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="ae"/>
      <w:pBdr>
        <w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="1" w:color="auto"/>
      </w:pBdr>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:hint="eastAsia"/>
      </w:rPr>
      <w:t>{text}</w:t>
    </w:r>
  </w:p>
</w:hdr>'''


def create_footer_xml():
    """创建带页码的页脚XML内容
    
    使用模板中的"af0"样式（页脚样式）
    """
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="af0"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:instrText xml:space="preserve"> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>'''


def extract_chapters(root):
    """从document.xml中提取章节信息（仅正文章节，用于公式编号等）
    
    注意：只遍历body直接子元素中的段落，不进入sdt等嵌套元素内部，
    避免将目录标题等误计入章节列表。
    """
    w_ns = NAMESPACES['w']
    chapters = []
    chapter_num = 0
    
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return chapters
    
    for child in body:
        if child.tag != f'{{{w_ns}}}p':
            continue
        para = child
        pPr = para.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue
        
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue
        
        style_val = pStyle.get(f'{{{w_ns}}}val', '')
        # 样式 "1" 或 "Heading1" 是一级标题
        if style_val in ('1', 'Heading1'):
            # 提取标题文本
            title_text = ''
            for t in para.findall(f'.//{{{w_ns}}}t'):
                if t.text:
                    title_text += t.text
            
            title_text = title_text.strip()
            # 跳过非正文章节（去除空格后比较，因为Quarto会压缩多空格）
            title_nospace = title_text.replace(' ', '').replace('\u3000', '')
            if any(exclude in title_nospace for exclude in EXCLUDE_CHAPTER_TITLES):
                continue
            
            chapter_num += 1
            chapters.append({
                'num': chapter_num,
                'title': title_text,
                'element': para
            })
    
    return chapters


def extract_all_headings(root):
    """从document.xml中提取所有一级标题（包括非正文章节）
    
    注意：只遍历body直接子元素中的段落，不进入sdt等嵌套元素内部，
    与update_sect_pr_references中的heading_counter保持一致，
    避免将目录标题等误计入导致页眉索引偏移。
    
    返回列表，每个元素包含：
    - title: 标题文本
    - is_numbered: 是否需要编号（正文章节为True）
    - chapter_num: 章节编号（仅正文章节有效）
    """
    w_ns = NAMESPACES['w']
    headings = []
    chapter_num = 0
    
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return headings
    
    for child in body:
        if child.tag != f'{{{w_ns}}}p':
            continue
        para = child
        pPr = para.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue
        
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue
        
        style_val = pStyle.get(f'{{{w_ns}}}val', '')
        # 样式 "1" 或 "Heading1" 是一级标题
        if style_val in ('1', 'Heading1'):
            # 提取标题文本
            title_text = ''
            for t in para.findall(f'.//{{{w_ns}}}t'):
                if t.text:
                    title_text += t.text
            
            title_text = title_text.strip()
            
            # 判断是否是非正文章节（去除空格后比较，因为Quarto会压缩多空格）
            title_nospace = title_text.replace(' ', '').replace('\u3000', '')
            is_excluded = any(exclude in title_nospace for exclude in EXCLUDE_CHAPTER_TITLES)
            
            if is_excluded:
                headings.append({
                    'title': title_text,
                    'is_numbered': False,
                    'chapter_num': None
                })
            else:
                chapter_num += 1
                headings.append({
                    'title': title_text,
                    'is_numbered': True,
                    'chapter_num': chapter_num
                })
    
    return headings


def enable_even_odd_headers(temp_dir):
    """在settings.xml中启用奇偶页不同页眉"""
    settings_path = temp_dir / "word" / "settings.xml"
    
    # 注册命名空间
    ET.register_namespace('w', NAMESPACES['w'])
    ET.register_namespace('r', NAMESPACES['r'])
    
    tree = ET.parse(settings_path)
    root = tree.getroot()
    
    w_ns = NAMESPACES['w']
    
    # 检查是否已有evenAndOddHeaders
    even_odd = root.find(f'{{{w_ns}}}evenAndOddHeaders')
    if even_odd is None:
        # 创建并添加
        even_odd = ET.Element(f'{{{w_ns}}}evenAndOddHeaders')
        # 插入到合适位置（通常在开头）
        root.insert(0, even_odd)
    
    tree.write(settings_path, encoding='utf-8', xml_declaration=True)
    print("  已启用奇偶页不同页眉设置")


def update_content_types(temp_dir, header_files, footer_files):
    """更新[Content_Types].xml"""
    ct_path = temp_dir / "[Content_Types].xml"
    
    tree = ET.parse(ct_path)
    root = tree.getroot()
    
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace('', ct_ns)
    
    # 检查现有的Override
    existing = set()
    for override in root.findall(f'{{{ct_ns}}}Override'):
        existing.add(override.get('PartName'))
    
    # 添加header文件
    for hf in header_files:
        part_name = f'/word/{hf}'
        if part_name not in existing:
            override = ET.SubElement(root, f'{{{ct_ns}}}Override')
            override.set('PartName', part_name)
            override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
    
    # 添加footer文件
    for ff in footer_files:
        part_name = f'/word/{ff}'
        if part_name not in existing:
            override = ET.SubElement(root, f'{{{ct_ns}}}Override')
            override.set('PartName', part_name)
            override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
    
    tree.write(ct_path, encoding='utf-8', xml_declaration=True)


def update_relationships(temp_dir, header_footer_map):
    """更新document.xml.rels"""
    rels_path = temp_dir / "word" / "_rels" / "document.xml.rels"
    
    rel_ns = NAMESPACES['rel']
    ET.register_namespace('', rel_ns)
    
    tree = ET.parse(rels_path)
    root = tree.getroot()
    
    # 找到最大的rId
    max_id = 0
    for rel in root:
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                num = int(rid[3:])
                max_id = max(max_id, num)
            except ValueError:
                pass
    
    # 添加新的relationships
    id_map = {}
    for filename, rel_type in header_footer_map.items():
        max_id += 1
        rid = f'rId{max_id}'
        id_map[filename] = rid
        
        rel = ET.SubElement(root, f'{{{rel_ns}}}Relationship')
        rel.set('Id', rid)
        rel.set('Type', rel_type)
        rel.set('Target', filename)
    
    tree.write(rels_path, encoding='utf-8', xml_declaration=True)
    return id_map


def apply_header_footer(temp_dir, root, thesis_title=None):
    """应用页眉页脚到文档"""
    print("正在处理页眉页脚...")
    
    w_ns = NAMESPACES['w']
    r_ns = NAMESPACES['r']
    word_dir = temp_dir / "word"
    
    # 1. 提取所有一级标题（包括非正文章节）
    all_headings = extract_all_headings(root)
    
    # 同时提取正文章节（用于显示信息）
    chapters = [h for h in all_headings if h['is_numbered']]
    print(f"  找到 {len(chapters)} 个正文章节，{len(all_headings) - len(chapters)} 个非正文章节")
    for h in all_headings:
        if h['is_numbered']:
            print(f"    第{h['chapter_num']}章: {h['title']}")
        else:
            print(f"    [非编号]: {h['title']}")
    
    # 2. 创建页眉页脚文件
    header_files = []
    footer_files = []
    header_footer_map = {}
    
    # 偶数页页眉（所有章节共用，包括目录部分）
    even_header_file = "header_even.xml"
    even_header_content = create_header_xml(EVEN_HEADER_TEXT)
    (word_dir / even_header_file).write_text(even_header_content, encoding='utf-8')
    header_files.append(even_header_file)
    header_footer_map[even_header_file] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    
    # 目录部分的奇数页页眉
    toc_header_file = "header_toc.xml"
    toc_header_text = thesis_title if thesis_title else TOC_TITLE
    toc_header_content = create_header_xml(toc_header_text)
    (word_dir / toc_header_file).write_text(toc_header_content, encoding='utf-8')
    header_files.append(toc_header_file)
    header_footer_map[toc_header_file] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    
    # 页脚（正文章节共用，目录部分不使用）
    footer_file = "footer_page.xml"
    footer_content = create_footer_xml()
    (word_dir / footer_file).write_text(footer_content, encoding='utf-8')
    footer_files.append(footer_file)
    header_footer_map[footer_file] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    
    # 为每个一级标题创建奇数页页眉
    heading_headers = {}  # 索引 -> 页眉文件名
    for i, h in enumerate(all_headings):
        odd_header_file = f"header_odd_{i+1}.xml"
        if h['is_numbered']:
            # 正文章节：第X章 标题
            odd_header_text = f"第{h['chapter_num']}章 {h['title']}"
        else:
            # 非正文章节：只有标题
            odd_header_text = h['title']
        
        odd_header_content = create_header_xml(odd_header_text)
        (word_dir / odd_header_file).write_text(odd_header_content, encoding='utf-8')
        header_files.append(odd_header_file)
        header_footer_map[odd_header_file] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
        heading_headers[i + 1] = odd_header_file  # sectPr索引从1开始（0是目录）
    
    print(f"  创建了 {len(header_files)} 个页眉文件和 {len(footer_files)} 个页脚文件")
    
    # 3. 更新Content_Types
    update_content_types(temp_dir, header_files, footer_files)
    
    # 4. 更新relationships并获取ID映射
    id_map = update_relationships(temp_dir, header_footer_map)
    
    # 5. 启用奇偶页不同
    enable_even_odd_headers(temp_dir)
    
    # 6. 更新document.xml中的sectPr
    update_sect_pr_references(root, all_headings, id_map, even_header_file, footer_file, heading_headers, toc_header_file)
    
    print("  页眉页脚处理完成")


def update_sect_pr_references(root, all_headings, id_map, even_header_file, footer_file, heading_headers, toc_header_file):
    """更新sectPr中的页眉页脚引用
    
    通过遍历body子元素动态识别每个sectPr的类型：
    - 'before_toc': 目录之前的sectPr（摘要/Abstract部分，有页眉，无页码）
    - 'toc': 目录后紧跟的sectPr（目录部分，有页眉，无页码）
    - 'content': 正文章节的sectPr（有页眉和页码）
    """
    w_ns = NAMESPACES['w']
    r_ns = NAMESPACES['r']
    
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        raise ValueError("找不到document body")
    
    children = list(body)
    
    # 找到sdt（目录）的位置
    sdt_index = None
    for i, child in enumerate(children):
        if child.tag.split('}')[-1] == 'sdt':
            sdt_index = i
            break
    
    # 遍历body子元素，收集所有sectPr及其类型
    # sect_info: (type, heading_counter, sectPr)
    sect_info_list = []
    heading_counter = 0
    found_toc_sect = False
    found_first_content_heading = False  # 是否已遇到第一个正文章节标题
    
    for i, child in enumerate(children):
        local_tag = child.tag.split('}')[-1]
        
        if local_tag == 'p':
            pPr = child.find(f'{{{w_ns}}}pPr')
            if pPr is None:
                continue
            
            # 检查是否是一级标题
            pStyle = pPr.find(f'{{{w_ns}}}pStyle')
            if pStyle is not None and pStyle.get(f'{{{w_ns}}}val', '') in (STYLES['heading1'], STYLES['heading1_alt']):
                heading_counter += 1
                # 检查是否为正文章节标题（不在排除列表中）
                if not found_first_content_heading and sdt_index is not None and i > sdt_index:
                    h_text = ''
                    for t in child.findall(f'.//{{{w_ns}}}t'):
                        if t.text:
                            h_text += t.text
                    h_text_nospace = h_text.strip().replace(' ', '').replace('\u3000', '')
                    is_excluded = any(exc in h_text_nospace for exc in EXCLUDE_CHAPTER_TITLES)
                    if not is_excluded:
                        found_first_content_heading = True
            
            # 检查是否包含sectPr
            sectPr = pPr.find(f'{{{w_ns}}}sectPr')
            if sectPr is not None:
                if sdt_index is not None and i < sdt_index:
                    sect_info_list.append(('before_toc', heading_counter, sectPr))
                elif sdt_index is not None and not found_toc_sect and i > sdt_index:
                    sect_info_list.append(('toc', 0, sectPr))
                    found_toc_sect = True
                elif found_toc_sect and not found_first_content_heading:
                    # 目录后、第一个正文章节前的sectPr（如创新成果自评表）
                    sect_info_list.append(('toc', 0, sectPr))
                else:
                    sect_info_list.append(('content', heading_counter, sectPr))
        
        elif local_tag == 'sectPr':
            # body末尾的sectPr
            sect_info_list.append(('content', heading_counter, child))
    
    print(f"  找到 {len(sect_info_list)} 个sectPr")
    for st, hc, _ in sect_info_list:
        print(f"    {st} (heading={hc})")
    
    # 处理每个sectPr
    content_sect_count = 0
    for sect_type, heading_idx, sectPr in sect_info_list:
        # 移除现有的headerReference和footerReference
        for ref in list(sectPr.findall(f'{{{w_ns}}}headerReference')):
            sectPr.remove(ref)
        for ref in list(sectPr.findall(f'{{{w_ns}}}footerReference')):
            sectPr.remove(ref)
        
        if sect_type == 'toc':
            # 目录部分：添加页眉，不添加页脚（不显示页码）
            toc_odd_ref = ET.Element(f'{{{w_ns}}}headerReference')
            toc_odd_ref.set(f'{{{w_ns}}}type', 'default')
            toc_odd_ref.set(f'{{{r_ns}}}id', id_map[toc_header_file])
            sectPr.insert(0, toc_odd_ref)
            
            toc_even_ref = ET.Element(f'{{{w_ns}}}headerReference')
            toc_even_ref.set(f'{{{w_ns}}}type', 'even')
            toc_even_ref.set(f'{{{r_ns}}}id', id_map[even_header_file])
            sectPr.insert(1, toc_even_ref)
            
            print("  目录部分sectPr：添加页眉，不添加页脚")
        
        elif sect_type == 'before_toc':
            # 摘要/Abstract部分：奇数页页眉用论文题目（与目录部分一致），不添加页脚
            odd_ref = ET.Element(f'{{{w_ns}}}headerReference')
            odd_ref.set(f'{{{w_ns}}}type', 'default')
            odd_ref.set(f'{{{r_ns}}}id', id_map[toc_header_file])
            sectPr.insert(0, odd_ref)
            
            even_ref = ET.Element(f'{{{w_ns}}}headerReference')
            even_ref.set(f'{{{w_ns}}}type', 'even')
            even_ref.set(f'{{{r_ns}}}id', id_map[even_header_file])
            sectPr.insert(1, even_ref)
            
            print(f"  摘要/Abstract sectPr (heading={heading_idx})：奇数页页眉用论文题目，不添加页脚")
        
        elif sect_type == 'content':
            content_sect_count += 1
            
            # 确定使用哪个奇数页页眉
            if heading_idx in heading_headers:
                odd_header = heading_headers[heading_idx]
            elif all_headings:
                odd_header = heading_headers[len(all_headings)]
            else:
                continue
            
            # 添加奇数页（default）页眉引用
            odd_ref = ET.Element(f'{{{w_ns}}}headerReference')
            odd_ref.set(f'{{{w_ns}}}type', 'default')
            odd_ref.set(f'{{{r_ns}}}id', id_map[odd_header])
            sectPr.insert(0, odd_ref)
            
            # 添加偶数页页眉引用
            even_ref = ET.Element(f'{{{w_ns}}}headerReference')
            even_ref.set(f'{{{w_ns}}}type', 'even')
            even_ref.set(f'{{{r_ns}}}id', id_map[even_header_file])
            sectPr.insert(1, even_ref)
            
            # 添加页脚引用（default和even都用同一个）
            footer_ref = ET.Element(f'{{{w_ns}}}footerReference')
            footer_ref.set(f'{{{w_ns}}}type', 'default')
            footer_ref.set(f'{{{r_ns}}}id', id_map[footer_file])
            sectPr.insert(2, footer_ref)
            
            footer_ref_even = ET.Element(f'{{{w_ns}}}footerReference')
            footer_ref_even.set(f'{{{w_ns}}}type', 'even')
            footer_ref_even.set(f'{{{r_ns}}}id', id_map[footer_file])
            sectPr.insert(3, footer_ref_even)
            
            # 第一个正文sectPr设置页码从1开始
            if content_sect_count == 1:
                for pgNum in list(sectPr.findall(f'{{{w_ns}}}pgNumType')):
                    sectPr.remove(pgNum)
                pgMar = sectPr.find(f'{{{w_ns}}}pgMar')
                if pgMar is not None:
                    pgMar_index = list(sectPr).index(pgMar)
                    pgNumType = ET.Element(f'{{{w_ns}}}pgNumType')
                    pgNumType.set(f'{{{w_ns}}}start', '1')
                    sectPr.insert(pgMar_index + 1, pgNumType)
                else:
                    pgNumType = ET.Element(f'{{{w_ns}}}pgNumType')
                    pgNumType.set(f'{{{w_ns}}}start', '1')
                    sectPr.append(pgNumType)
                print("  第1章sectPr：设置页码从1开始")
