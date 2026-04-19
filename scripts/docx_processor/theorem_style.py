"""
定理环境样式处理模块
处理定理、定义、引理等环境的交叉引用
"""
import xml.etree.ElementTree as ET
import re
from .utils import NAMESPACES


# 定理环境类型映射
THEOREM_TYPES = {
    'def': '定义',
    'thm': '定理',
    'lem': '引理',
    'cor': '推论',
    'prp': '命题',
    'exm': '例',
    'rem': '注',
}


def get_paragraph_text(para):
    """获取段落的纯文本内容"""
    w_ns = NAMESPACES['w']
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


def build_theorem_map(root):
    """构建定理环境书签到显示文本的映射
    
    扫描文档中的书签，识别定理环境类型，并提取显示文本
    """
    w_ns = NAMESPACES['w']
    theorem_map = {}
    
    # 查找所有定理环境书签
    for bookmark in root.iter(f'{{{w_ns}}}bookmarkStart'):
        name = bookmark.get(f'{{{w_ns}}}name', '')
        
        # 检查是否是定理环境书签
        for prefix, cn_name in THEOREM_TYPES.items():
            if name.startswith(f'{prefix}-'):
                # 找到包含此书签的段落
                parent = find_parent(root, bookmark)
                while parent is not None and parent.tag != f'{{{w_ns}}}p':
                    parent = find_parent(root, parent)
                
                if parent is not None:
                    text = get_paragraph_text(parent)
                    # 提取定理编号，如 "定义 1.1（凸函数）" -> "定义 1.1"
                    match = re.search(rf'{cn_name}\s*(\d+\.\d+)', text)
                    if match:
                        display_text = f'{cn_name}{match.group(1)}'
                        theorem_map[name] = display_text
                break
    
    return theorem_map


def _make_text_run(w_ns, text, rPr_template=None):
    """创建一个携带文本的 w:r 元素，保留原有 rPr（若提供）。"""
    r = ET.Element(f'{{{w_ns}}}r')
    if rPr_template is not None:
        # 浅拷贝 rPr 以保留原有字体/字号等属性
        r.append(ET.fromstring(ET.tostring(rPr_template)))
    t = ET.SubElement(r, f'{{{w_ns}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return r


def _make_hyperlink_run(w_ns, ref_id, display_text):
    """创建一个指向 ref_id 的超链接元素，内含 Hyperlink 样式的 run。"""
    hyperlink = ET.Element(f'{{{w_ns}}}hyperlink')
    hyperlink.set(f'{{{w_ns}}}anchor', ref_id)
    new_run = ET.Element(f'{{{w_ns}}}r')
    new_rPr = ET.SubElement(new_run, f'{{{w_ns}}}rPr')
    rStyle = ET.SubElement(new_rPr, f'{{{w_ns}}}rStyle')
    rStyle.set(f'{{{w_ns}}}val', 'Hyperlink')
    new_t = ET.SubElement(new_run, f'{{{w_ns}}}t')
    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    new_t.text = display_text
    hyperlink.append(new_run)
    return hyperlink


def update_theorem_references(root, theorem_map):
    """更新定理环境的交叉引用

    将 run 文本中的 "?@def-xxx" 片段拆分并替换为超链接：
    原 run: [前文本 ?@lem-a 中间文本 ?@thm-b 后文本]
    拆分后: [前文本 run][hyperlink(引理 3.1)][中间文本 run][hyperlink(定理 3.1)][后文本 run]

    这样可保留原文本中超链接前后的空格为独立 run，便于后续 strip_spaces_around_cross_refs 处理。
    """
    w_ns = NAMESPACES['w']
    updated_count = 0
    pattern = re.compile(r'\?@([a-zA-Z]+-[a-zA-Z0-9_-]+)')

    for para in root.iter(f'{{{w_ns}}}p'):
        text = get_paragraph_text(para)
        if '?@' not in text:
            continue

        for run in list(para.iter(f'{{{w_ns}}}r')):
            t_elem = run.find(f'{{{w_ns}}}t')
            if t_elem is None or t_elem.text is None:
                continue
            original_text = t_elem.text
            if '?@' not in original_text:
                continue

            # 收集所有命中定理映射的匹配
            matches = [m for m in pattern.finditer(original_text) if m.group(1) in theorem_map]
            if not matches:
                continue

            rPr_template = run.find(f'{{{w_ns}}}rPr')
            parent = find_parent(para, run)
            if parent is None:
                continue
            insert_idx = list(parent).index(run)

            # 构造替换序列：前文本 run、hyperlink、中间文本 run、... 后文本 run
            new_elements = []
            cursor = 0
            for m in matches:
                start, end = m.start(), m.end()
                if start > cursor:
                    new_elements.append(_make_text_run(w_ns, original_text[cursor:start], rPr_template))
                new_elements.append(_make_hyperlink_run(w_ns, m.group(1), theorem_map[m.group(1)]))
                updated_count += 1
                cursor = end
            if cursor < len(original_text):
                new_elements.append(_make_text_run(w_ns, original_text[cursor:], rPr_template))

            # 移除原 run，按顺序插入新元素
            parent.remove(run)
            for offset, elem in enumerate(new_elements):
                parent.insert(insert_idx + offset, elem)

    return updated_count


def process_theorem_references(root):
    """处理定理环境的交叉引用"""
    print("正在处理定理环境交叉引用...")
    
    # 构建定理映射
    theorem_map = build_theorem_map(root)
    print(f"  找到 {len(theorem_map)} 个定理环境书签")
    for name, display in theorem_map.items():
        print(f"    {name} -> {display}")
    
    # 更新交叉引用
    updated_count = update_theorem_references(root, theorem_map)
    print(f"  更新了 {updated_count} 个定理交叉引用")
    
    return theorem_map
