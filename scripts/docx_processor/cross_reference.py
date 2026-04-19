import xml.etree.ElementTree as ET
import re
from .utils import NAMESPACES
from .config import EXCLUDE_CHAPTER_TITLES


# 标题样式ID到级别的映射（兼容不同模板）
_HEADING_STYLES = {
    '1': 1, 'Heading1': 1, 'heading1': 1,
    '2': 2, 'Heading2': 2, 'heading2': 2,
    '3': 3, 'Heading3': 3, 'heading3': 3,
    '4': 4, 'Heading4': 4, 'heading4': 4,
}


def _get_paragraph_text(para):
    """获取段落的纯文本内容"""
    w_ns = NAMESPACES['w']
    text = ''
    for t in para.findall(f'.//{{{w_ns}}}t'):
        if t.text:
            text += t.text
    return text.strip()


def _build_section_map(root):
    """构建章节书签到编号的映射

    扫描文档中的标题段落，按层级计算编号，并建立 sec-xxx 书签名到编号字符串的映射。
    排除非正文章节（摘要、结论、参考文献、致谢等）。

    注意：Quarto 生成的 docx 中，sec- 书签是 body 的直接子元素，
    位于对应标题段落之前（而非段落内部），因此需要遍历 body 子元素，
    先收集待处理书签，遇到标题段落时再进行关联。
    """
    w_ns = NAMESPACES['w']
    section_map = {}  # 书签名 -> (级别, 编号字符串)
    counters = [0, 0, 0, 0]  # 四级计数器
    in_excluded = False  # 当前是否处于被排除的一级章节中

    # 查找 body 元素
    body = root.find(f'{{{w_ns}}}body')
    if body is None:
        body = root

    # 遍历 body 的直接子元素
    pending_sec_bookmarks = []  # 暂存待关联的 sec- 书签名
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        # 收集 sec- 书签（body 直接子元素）
        if tag == 'bookmarkStart':
            name = child.get(f'{{{w_ns}}}name', '')
            if name.startswith('sec-'):
                pending_sec_bookmarks.append(name)
            continue

        # 处理段落
        if tag != 'p':
            continue

        pPr = child.find(f'{{{w_ns}}}pPr')
        if pPr is None:
            continue

        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is None:
            continue

        style_val = pStyle.get(f'{{{w_ns}}}val', '')
        level = _HEADING_STYLES.get(style_val)
        if level is None:
            continue

        title_text = _get_paragraph_text(child)

        # 一级标题：检查是否为非正文章节
        if level == 1:
            title_text_nospace = title_text.replace(' ', '').replace('\u3000', '')
            is_excluded = any(kw in title_text_nospace for kw in EXCLUDE_CHAPTER_TITLES)
            in_excluded = is_excluded
            if is_excluded:
                pending_sec_bookmarks.clear()
                continue
            # 更新一级计数器，重置下级
            counters[0] += 1
            for i in range(1, len(counters)):
                counters[i] = 0
        else:
            # 非一级标题：如果当前处于被排除的章节中，跳过
            if in_excluded:
                pending_sec_bookmarks.clear()
                continue
            # 更新当前级别计数器，重置下级
            counters[level - 1] += 1
            for i in range(level, len(counters)):
                counters[i] = 0

        # 生成编号字符串，如 "2" (一级) 或 "2.3" (二级)
        number_str = '.'.join(str(counters[i]) for i in range(level))

        # 将之前收集的 sec- 书签与当前标题关联
        for bm_name in pending_sec_bookmarks:
            section_map[bm_name] = (level, number_str)
        pending_sec_bookmarks.clear()

        # 同时检查段落内部的书签（兼容不同 Quarto 版本）
        for bookmark in child.findall(f'.//{{{w_ns}}}bookmarkStart'):
            name = bookmark.get(f'{{{w_ns}}}name', '')
            if name.startswith('sec-') and name not in section_map:
                section_map[name] = (level, number_str)

    return section_map


def update_cross_references(root, equation_map):
    """更新交叉引用格式：公式引用为"公式（章-序号）"，章节引用为"第X章"/"第X.Y节"

    公式引用会自动加"公式"前缀，并移除正文中已有的"公式"/"式"前缀避免重复。
    与图引用（自动加"图"）、表引用（自动加"表"）的处理逻辑保持统一。
    """
    print("开始更新交叉引用...")
    refs_count = 0
    sec_refs_count = 0
    w_ns = NAMESPACES['w']

    # 构建章节编号映射
    section_map = _build_section_map(root)
    if section_map:
        print(f"  章节书签映射:")
        for name, (lvl, num) in section_map.items():
            print(f"    {name} -> 级别{lvl}, 编号{num}")

    for para in root.findall('.//w:p', NAMESPACES):
        children = list(para)
        to_remove = []

        for i, child in enumerate(children):
            if child.tag != f'{{{w_ns}}}hyperlink':
                continue
            anchor = child.get(f'{{{w_ns}}}anchor')

            # 处理章节引用
            if anchor and anchor.startswith('sec-'):
                runs = child.findall('w:r', NAMESPACES)
                if anchor in section_map:
                    level, number_str = section_map[anchor]
                    # 一级标题 -> "第 X 章"，其余 -> "第 X.Y 节"
                    if level == 1:
                        display_text = f'第{number_str}章'
                    else:
                        display_text = f'第{number_str}节'
                    # 将第一个 run 的文本替换，删除多余 run
                    if runs:
                        first_run = runs[0]
                        t_elem = first_run.find('w:t', NAMESPACES)
                        if t_elem is None:
                            t_elem = ET.Element(f'{{{w_ns}}}t')
                            first_run.append(t_elem)
                        t_elem.text = display_text
                        # 移除Hyperlink字符样式，避免与正文产生视觉间距差异
                        rPr = first_run.find('w:rPr', NAMESPACES)
                        if rPr is not None:
                            rStyle = rPr.find('w:rStyle', NAMESPACES)
                            if rStyle is not None:
                                rPr.remove(rStyle)
                        for r in runs[1:]:
                            child.remove(r)
                        sec_refs_count += 1

                    # 移除前面的冗余"第X章"/"第X节"文本
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
                                # 移除末尾的"第X章"或"第X节"
                                new_text = re.sub(
                                    r'第[一二三四五六七八九十\d]+[章节]\s*$', '', text)
                                if new_text != text:
                                    if new_text.strip():
                                        t.text = new_text
                                    else:
                                        to_remove.append(prev)
                                    break
                                else:
                                    break
                            else:
                                break
                        else:
                            break
                        j -= 1

                    # 移除后面的冗余"节"/"章"文本
                    j = i + 1
                    while j < len(children):
                        nxt = children[j]
                        if nxt.tag == f'{{{w_ns}}}r':
                            t = nxt.find(f'{{{w_ns}}}t')
                            if t is not None and t.text is not None:
                                text = t.text
                                if text.strip() == '':
                                    j += 1
                                    continue
                                # 移除开头的"节"或"章"（允许前导空格）
                                new_text = re.sub(r'^\s*[节章]\s*', '', text)
                                if new_text != text:
                                    if new_text.strip():
                                        t.text = new_text
                                    else:
                                        to_remove.append(nxt)
                                    break
                                else:
                                    break
                            else:
                                break
                        else:
                            break
                        j += 1

                else:
                    # 回退：简单提取数字
                    for run in runs:
                        t_elem = run.find('w:t', NAMESPACES)
                        if t_elem is not None and t_elem.text:
                            match = re.match(r'Section\s+([\d.]+)', t_elem.text)
                            if match:
                                t_elem.text = match.group(1)
                                sec_refs_count += 1

            if anchor and anchor in equation_map:
                # 更新引用文本为"公式（章-序号）"格式
                chap_num, eq_num = equation_map[anchor]
                new_text = f"公式（{chap_num}-{eq_num}）"

                runs = child.findall('w:r', NAMESPACES)
                if runs:
                    first_run = runs[0]
                    t_elem = first_run.find('w:t', NAMESPACES)
                    if t_elem is None:
                        t_elem = ET.Element(f'{{{w_ns}}}t')
                        first_run.append(t_elem)

                    t_elem.text = new_text
                    # 移除Hyperlink字符样式，避免与正文产生视觉间距差异
                    rPr = first_run.find('w:rPr', NAMESPACES)
                    if rPr is not None:
                        rStyle = rPr.find('w:rStyle', NAMESPACES)
                        if rStyle is not None:
                            rPr.remove(rStyle)

                    # 删除其他的runs
                    for r in runs[1:]:
                        child.remove(r)

                    refs_count += 1

                    # 移除前面已有的"公式"/"式"前缀，避免重复
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
                                # 移除末尾的"公式"或"式"
                                new_t = re.sub(r'(公式|式)\s*$', '', text)
                                if new_t != text:
                                    if new_t.strip():
                                        t.text = new_t
                                    else:
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

    print(f"成功更新 {refs_count} 个公式交叉引用，{sec_refs_count} 个章节交叉引用")


def strip_spaces_around_cross_refs(root):
    """移除交叉引用超链接前后的多余空格

    Quarto 编译时，QMD 中 `图 @fig-xxx 所示` 的空格会被保留到 docx 中，
    导致中文文本与交叉引用之间出现不必要的空格。
    本函数遍历所有段落，移除带 w:anchor 属性的超链接前后文本 run 中的
    尾部/头部空格。
    """
    w_ns = NAMESPACES['w']
    count = 0

    for para in root.iter(f'{{{w_ns}}}p'):
        children = list(para)

        for i, child in enumerate(children):
            if child.tag != f'{{{w_ns}}}hyperlink':
                continue
            # 仅处理内部交叉引用（带 anchor 属性）
            anchor = child.get(f'{{{w_ns}}}anchor')
            if not anchor:
                continue

            # 移除前一个 run 末尾的空格
            if i > 0:
                prev = children[i - 1]
                if prev.tag == f'{{{w_ns}}}r':
                    t = prev.find(f'{{{w_ns}}}t')
                    if t is not None and t.text and t.text.endswith(' '):
                        t.text = t.text.rstrip(' ')
                        count += 1

            # 移除后一个 run 开头的空格
            if i + 1 < len(children):
                nxt = children[i + 1]
                if nxt.tag == f'{{{w_ns}}}r':
                    t = nxt.find(f'{{{w_ns}}}t')
                    if t is not None and t.text and t.text.startswith(' '):
                        t.text = t.text.lstrip(' ')
                        count += 1

    print(f"  移除了 {count} 处交叉引用前后的多余空格")
