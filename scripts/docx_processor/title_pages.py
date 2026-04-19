"""
封面页合并模块
功能：将封面页docx（thesis_title_pages.docx）的内容合并到论文最前面

分两步执行：
1. merge_title_pages_content(root, ...) - 在save_document_xml之前调用，修改内存中的root
2. merge_title_pages_files(temp_path, ...) - 在save_document_xml之后调用，复制文件和修改rels
"""
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile
import shutil

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# bookmark ID偏移量，避免与thesis中的ID冲突
BOOKMARK_ID_OFFSET = 10000


def _remove_bookmarks(elements, w_ns):
    """移除封面页中所有bookmark（模板遗留的_Toc书签，与thesis中name冲突）
    
    返回需要从elements列表中移除的body直接子元素（bookmarkStart/End）
    """
    to_remove = []
    for elem in elements:
        local_tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        # body直接子元素中的bookmarkStart/End需要整个移除
        if local_tag in ('bookmarkStart', 'bookmarkEnd'):
            to_remove.append(elem)
            continue
        # 段落等元素内部的bookmark需要从父元素中移除
        for bm in list(elem.findall(f'.//{{{w_ns}}}bookmarkStart')):
            parent = _find_parent(elem, bm, w_ns)
            if parent is not None:
                parent.remove(bm)
        for bm in list(elem.findall(f'.//{{{w_ns}}}bookmarkEnd')):
            parent = _find_parent(elem, bm, w_ns)
            if parent is not None:
                parent.remove(bm)
    return to_remove


def _find_parent(root, target, w_ns):
    """在root的子树中找到target的直接父元素"""
    for parent in root.iter():
        for child in parent:
            if child is target:
                return parent
    return None


def prepare_title_pages(title_pages_docx_path):
    """解压封面页docx并提取内容和关系信息
    
    Returns:
        dict with keys: content_elements, body_sectPr, hf_map, tp_temp_path
        或 None（如果文件不存在）
    """
    title_pages_path = Path(title_pages_docx_path)
    if not title_pages_path.exists():
        print(f"  封面页文件不存在: {title_pages_path}，跳过合并")
        return None
    
    w_ns = NAMESPACES['w']
    
    # 解压到临时目录（调用者负责清理）
    import tempfile
    tp_temp = Path(tempfile.mkdtemp(prefix='tp_'))
    with ZipFile(title_pages_path, 'r') as zf:
        zf.extractall(tp_temp)
    
    # 读取document.xml
    tp_doc_path = tp_temp / "word" / "document.xml"
    tp_tree = ET.parse(tp_doc_path)
    tp_root = tp_tree.getroot()
    tp_body = tp_root.find(f'{{{w_ns}}}body')
    if tp_body is None:
        shutil.rmtree(tp_temp)
        return None
    
    # 分离body-level sectPr和其他子元素
    tp_body_sectPr = None
    tp_content_elements = []
    for child in list(tp_body):
        if child.tag == f'{{{w_ns}}}sectPr':
            tp_body_sectPr = child
        else:
            tp_content_elements.append(child)
    
    if not tp_content_elements:
        shutil.rmtree(tp_temp)
        return None
    
    # 读取rels，提取header/footer关系
    tp_rels_path = tp_temp / "word" / "_rels" / "document.xml.rels"
    tp_hf_map = {}
    if tp_rels_path.exists():
        tp_rels_tree = ET.parse(tp_rels_path)
        for rel in tp_rels_tree.getroot():
            rel_type = rel.get('Type', '')
            rel_id = rel.get('Id', '')
            rel_target = rel.get('Target', '')
            if 'header' in rel_type or 'footer' in rel_type:
                tp_hf_map[rel_id] = (rel_type, rel_target)
    
    return {
        'content_elements': tp_content_elements,
        'body_sectPr': tp_body_sectPr,
        'hf_map': tp_hf_map,
        'tp_temp_path': tp_temp,
    }


def merge_title_pages_content(root, tp_data, new_rid_map):
    """将封面页内容插入thesis body最前面（操作内存中的root）
    
    在save_document_xml之前调用。
    
    Args:
        root: thesis document.xml的根元素
        tp_data: prepare_title_pages返回的数据
        new_rid_map: 旧rId -> 新rId的映射（由merge_title_pages_files预先计算）
    """
    w_ns = NAMESPACES['w']
    r_ns = NAMESPACES['r']
    
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        return
    
    tp_content = tp_data['content_elements']
    tp_sectPr = tp_data['body_sectPr']
    
    # 移除封面页中所有bookmark（模板遗留_Toc书签，与thesis中name冲突）
    to_remove = _remove_bookmarks(tp_content, w_ns)
    for elem in to_remove:
        tp_content.remove(elem)
    if tp_sectPr is not None:
        _remove_bookmarks([tp_sectPr], w_ns)
    
    # 构建分节符段落
    if tp_sectPr is not None:
        # 更新header/footer rId引用
        for href in tp_sectPr.findall(f'{{{w_ns}}}headerReference'):
            old_rid = href.get(f'{{{r_ns}}}id', '')
            if old_rid in new_rid_map:
                href.set(f'{{{r_ns}}}id', new_rid_map[old_rid])
        for fref in tp_sectPr.findall(f'{{{w_ns}}}footerReference'):
            old_rid = fref.get(f'{{{r_ns}}}id', '')
            if old_rid in new_rid_map:
                fref.set(f'{{{r_ns}}}id', new_rid_map[old_rid])
        
        # 添加分节符类型
        # CT_SectPr 子元素顺序：headerReference* → footerReference* →
        #   endnotePr? → footnotePr? → type? → pgSz? → ...
        # type 必须在 headerReference/footerReference/endnotePr/footnotePr 之后
        sect_type = tp_sectPr.find(f'{{{w_ns}}}type')
        if sect_type is None:
            sect_type = ET.Element(f'{{{w_ns}}}type')
            _before_type = {'headerReference', 'footerReference',
                            'endnotePr', 'footnotePr'}
            insert_idx = 0
            for i, child in enumerate(tp_sectPr):
                if child.tag.split('}')[-1] in _before_type:
                    insert_idx = i + 1
            tp_sectPr.insert(insert_idx, sect_type)
        sect_type.set(f'{{{w_ns}}}val', 'nextPage')
        
        sect_break_para = ET.Element(f'{{{w_ns}}}p')
        sect_break_pPr = ET.SubElement(sect_break_para, f'{{{w_ns}}}pPr')
        sect_break_pPr.append(tp_sectPr)
    else:
        sect_break_para = ET.Element(f'{{{w_ns}}}p')
        sect_break_pPr = ET.SubElement(sect_break_para, f'{{{w_ns}}}pPr')
        new_sectPr = ET.SubElement(sect_break_pPr, f'{{{w_ns}}}sectPr')
        sect_type = ET.SubElement(new_sectPr, f'{{{w_ns}}}type')
        sect_type.set(f'{{{w_ns}}}val', 'nextPage')
    
    # 插入到body最前面
    insert_index = 0
    for elem in tp_content:
        body.insert(insert_index, elem)
        insert_index += 1
    body.insert(insert_index, sect_break_para)
    
    print(f"  已插入封面页内容（{len(tp_content)}个元素 + 分节符）")


def compute_rid_map(temp_path, tp_data):
    """预计算rId映射（不修改文件，只读取rels获取最大rId）
    
    在save_document_xml之前调用，以便merge_title_pages_content使用。
    
    Returns:
        dict: 旧rId -> 新rId的映射
    """
    thesis_rels_path = temp_path / "word" / "_rels" / "document.xml.rels"
    tp_hf_map = tp_data['hf_map']
    
    # 读取rels找最大rId（只读，不修改）
    max_rid_num = 0
    if thesis_rels_path.exists():
        thesis_rels_tree = ET.parse(thesis_rels_path)
        for rel in thesis_rels_tree.getroot():
            rid = rel.get('Id', '')
            if rid.startswith('rId'):
                try:
                    max_rid_num = max(max_rid_num, int(rid[3:]))
                except ValueError:
                    pass
    
    old_rid_to_new = {}
    for old_rid in tp_hf_map:
        max_rid_num += 1
        old_rid_to_new[old_rid] = f"rId{max_rid_num}"
    
    return old_rid_to_new


def merge_title_pages_files(temp_path, tp_data, rid_map):
    """复制header/footer文件、更新rels和Content_Types
    
    在save_document_xml之后、pack_docx之前调用。
    
    Args:
        temp_path: thesis解压后的临时目录
        tp_data: prepare_title_pages返回的数据
        rid_map: compute_rid_map返回的旧rId->新rId映射
    """
    thesis_word_dir = temp_path / "word"
    tp_temp = tp_data['tp_temp_path']
    tp_hf_map = tp_data['hf_map']
    
    # 1. 复制header/footer文件
    tp_file_rename = {}
    for rid, (rtype, target) in tp_hf_map.items():
        old_filename = target
        new_filename = f"tp_{old_filename}"
        src = tp_temp / "word" / old_filename
        dst = thesis_word_dir / new_filename
        if src.exists():
            shutil.copy2(src, dst)
            tp_file_rename[old_filename] = new_filename
    
    # 2. 更新[Content_Types].xml
    content_types_path = temp_path / "[Content_Types].xml"
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    ET.register_namespace('', ct_ns)
    ct_tree = ET.parse(content_types_path)
    ct_root = ct_tree.getroot()
    
    HEADER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
    FOOTER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
    
    for old_filename, new_filename in tp_file_rename.items():
        part_name = f"/word/{new_filename}"
        ct = HEADER_CT if 'header' in new_filename else FOOTER_CT
        override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
        override.set('PartName', part_name)
        override.set('ContentType', ct)
    
    ct_tree.write(content_types_path, encoding='utf-8', xml_declaration=True)
    
    # 3. 更新document.xml.rels
    rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ET.register_namespace('', rels_ns)
    thesis_rels_path = thesis_word_dir / "_rels" / "document.xml.rels"
    thesis_rels_tree = ET.parse(thesis_rels_path)
    thesis_rels_root = thesis_rels_tree.getroot()
    
    for old_rid, (rtype, target) in tp_hf_map.items():
        new_rid = rid_map[old_rid]
        new_target = tp_file_rename.get(target, f"tp_{target}")
        new_rel = ET.SubElement(thesis_rels_root, f'{{{rels_ns}}}Relationship')
        new_rel.set('Id', new_rid)
        new_rel.set('Type', rtype)
        new_rel.set('Target', new_target)
    
    thesis_rels_tree.write(thesis_rels_path, encoding='utf-8', xml_declaration=True)
    
    # 4. 清理临时目录
    if tp_temp.exists():
        shutil.rmtree(tp_temp)
    
    print(f"  复制了{len(tp_file_rename)}个header/footer文件，分配了{len(rid_map)}个新rId")
