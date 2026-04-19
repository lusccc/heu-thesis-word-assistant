"""
脚注样式处理模块

功能：
1. 在 settings.xml 中添加全局 footnotePr 设置
   - 编号格式：①②③ (decimalEnclosedCircle)
   - 每页重新编号 (numRestart=eachPage)
   - 位置：页面底部 (pageBottom)
2. 更新脚注文字样式（小五号宋体/Times New Roman）
3. 更新脚注引用样式（小五号上标）

注：Pandoc/Quarto 直接生成 footnotes.xml（非 endnotes），
因此无需尾注→脚注转换。

重要：w:footnotePr 在 OOXML schema 中属于 w:settings 子元素，
不能放在 w:styles 下，否则 Word 打开时会提示
“发现无法读取的内容”。
"""
import xml.etree.ElementTree as ET
from pathlib import Path
from .utils import fix_mc_ignorable_namespaces

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# 脚注编号格式：①②③
FOOTNOTE_NUM_FMT = 'decimalEnclosedCircle'
# 脚注文字大小：小五号 = 9pt = 18 半磅
FOOTNOTE_FONT_SIZE = '18'

# Pandoc/Quarto 生成的脚注样式 ID
PANDOC_FOOTNOTE_TEXT_STYLE = 'afa'
PANDOC_FOOTNOTE_REF_STYLE = 'afc'


def process_footnotes(temp_path, root):
    """设置脚注样式：①②③编号、每页重启、小五号字体

    Args:
        temp_path: 解压后的docx临时目录
        root: document.xml的根元素
    """
    footnotes_path = temp_path / "word" / "footnotes.xml"
    styles_path = temp_path / "word" / "styles.xml"
    settings_path = temp_path / "word" / "settings.xml"

    # 1. settings.xml 全局设置（仅 numFmt 等属性，不带 footnote 引用）
    _update_settings_xml(settings_path)
    # 2. styles.xml 脚注文字/引用样式
    _update_styles_xml(styles_path)
    # 3. document.xml 每个 sectPr 注入 footnotePr（节级设置最可靠）
    _inject_sectpr_footnote_pr(root)

    if not footnotes_path.exists():
        print("  未找到脚注(footnotes.xml)，仅写入脚注全局设置")
        return

    fn_count = _count_footnotes(footnotes_path)
    if fn_count == 0:
        print("  没有脚注内容，仅写入脚注全局设置")
        return

    print(f"  找到 {fn_count} 个脚注，脚注样式已设置（①②③编号，每页重启，小五号字体）")


def _count_footnotes(footnotes_path):
    """统计脚注数量（排除分隔符条目）"""
    try:
        tree = ET.parse(footnotes_path)
        fn_root = tree.getroot()
        count = 0
        for fn in fn_root:
            fn_type = fn.get(f'{{{W}}}type')
            if fn_type in ('separator', 'continuationSeparator'):
                continue
            count += 1
        return count
    except ET.ParseError:
        return 0


def _update_styles_xml(styles_path):
    """在 styles.xml 中更新脚注文字 / 脚注引用样式。

    注意：w:footnotePr 不属于 w:styles，已移至 settings.xml。
    """
    if not styles_path.exists():
        return

    tree = ET.parse(styles_path)
    styles_root = tree.getroot()

    # 如历史生成物在 styles.xml 根下存在非法的 footnotePr，顺手清理掉
    stale_fnpr = styles_root.find(f'{{{W}}}footnotePr')
    if stale_fnpr is not None:
        styles_root.remove(stale_fnpr)

    # 更新脚注文字样式 (afa = footnote text)
    _update_footnote_text_style(styles_root)

    # 更新脚注引用样式 (afc = footnote reference)
    _update_footnote_ref_style(styles_root)

    tree.write(styles_path, encoding='utf-8', xml_declaration=True)
    fix_mc_ignorable_namespaces(styles_path)


def _build_footnote_pr_element():
    """构建一个 footnotePr 元素（仅含 numFmt 属性，不含 footnote 引用）。

    子元素严格按 CT_FtnDocProps 顺序：pos → numFmt → numStart → numRestart。
    不包含 <w:footnote> separator 引用，避免 Word schema 验证失败。
    """
    fnpr = ET.Element(f'{{{W}}}footnotePr')
    pos = ET.SubElement(fnpr, f'{{{W}}}pos')
    pos.set(f'{{{W}}}val', 'pageBottom')
    numfmt = ET.SubElement(fnpr, f'{{{W}}}numFmt')
    numfmt.set(f'{{{W}}}val', FOOTNOTE_NUM_FMT)
    numrestart = ET.SubElement(fnpr, f'{{{W}}}numRestart')
    numrestart.set(f'{{{W}}}val', 'eachPage')
    return fnpr


_SECTPR_BEFORE_FOOTNOTEPR = frozenset([
    'headerReference', 'footerReference', 'endnotePr',
])

_SECTPR_AFTER_FOOTNOTEPR = frozenset([
    'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
    'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
    'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings',
    'sectPrChange',
])


def _inject_sectpr_footnote_pr(root):
    """在 document.xml 每个 sectPr 中添加/替换 footnotePr。

    节级 (sectPr) 的 footnotePr 是 Word 实际渲染脚注编号格式的最可靠来源。

    CT_SectPr 子元素严格顺序：
      headerReference* → footerReference* → endnotePr? → footnotePr?
      → type? → pgSz? → pgMar? → …
    必须插入到正确位置，否则 Word 报"无法读取的内容"。
    """
    before_tags = {f'{{{W}}}{n}' for n in _SECTPR_BEFORE_FOOTNOTEPR}
    count = 0
    for sectpr in root.iter(f'{{{W}}}sectPr'):
        # 移除已有的 footnotePr
        existing = sectpr.find(f'{{{W}}}footnotePr')
        if existing is not None:
            sectpr.remove(existing)

        # 找到最后一个"应在 footnotePr 之前"的元素，在其后插入
        insert_idx = 0
        for i, child in enumerate(sectpr):
            if child.tag in before_tags:
                insert_idx = i + 1

        fnpr = _build_footnote_pr_element()
        sectpr.insert(insert_idx, fnpr)
        count += 1
    print(f"  已在 {count} 个 sectPr 中注入 footnotePr（①②③，每页重启）")


# CT_Settings 中 footnotePr 后面常见的元素：按 schema 顺序作为插入锚点
# （footnotePr 必须位于这些元素之前）
_SETTINGS_AFTER_FOOTNOTEPR = (
    'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr',
    'attachedSchema', 'themeFontLang', 'clrSchemeMapping',
    'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures',
    'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType',
    'schemaLibrary', 'shapeDefaults', 'decimalSymbol', 'listSeparator',
)


def _update_settings_xml(settings_path):
    """在 settings.xml 的 w:settings 下添加/更新 w:footnotePr。

    按 OOXML schema (CT_Settings) 要求，w:footnotePr 必须出现在
    w:endnotePr / w:compat / w:rsids 等元素之前。若这些锚点都不存在，
    则追加到末尾。
    """
    if not settings_path.exists():
        print("  警告：未找到 settings.xml，跳过全局 footnotePr 设置")
        return

    tree = ET.parse(settings_path)
    settings_root = tree.getroot()

    # 清理已存在的 footnotePr
    existing_fnpr = settings_root.find(f'{{{W}}}footnotePr')
    if existing_fnpr is not None:
        settings_root.remove(existing_fnpr)

    fnpr = _build_footnote_pr_element()
    # 注意：settings.xml 级 footnotePr 不应包含 <w:footnote> 子元素引用，
    # 否则 Word 打开时会提示"无法读取的内容"。

    # 找到第一个“应在 footnotePr 之后”的锚点，在其前插入
    insert_idx = len(list(settings_root))
    anchor_tags = {f'{{{W}}}{name}' for name in _SETTINGS_AFTER_FOOTNOTEPR}
    for i, child in enumerate(settings_root):
        if child.tag in anchor_tags:
            insert_idx = i
            break

    settings_root.insert(insert_idx, fnpr)
    tree.write(settings_path, encoding='utf-8', xml_declaration=True)
    fix_mc_ignorable_namespaces(settings_path)


def _update_footnote_text_style(styles_root):
    """更新脚注文字样式：添加小五号字体大小"""
    for style_elem in styles_root.findall(f'{{{W}}}style'):
        style_id = style_elem.get(f'{{{W}}}styleId')
        style_name = ''
        name_elem = style_elem.find(f'{{{W}}}name')
        if name_elem is not None:
            style_name = name_elem.get(f'{{{W}}}val', '')

        # 匹配 Pandoc 生成的脚注文字样式或标准 FootnoteText 样式
        if style_id in (PANDOC_FOOTNOTE_TEXT_STYLE, 'FootnoteText') or style_name == 'footnote text':
            _ensure_font_size(style_elem, FOOTNOTE_FONT_SIZE)
            # 确保间距设置合理
            _ensure_paragraph_spacing(style_elem, after='0', line='240', line_rule='auto')
            return

    # 如果没有找到，创建新样式
    _create_footnote_text_style(styles_root)


def _update_footnote_ref_style(styles_root):
    """更新脚注引用样式：添加小五号字体大小和上标"""
    for style_elem in styles_root.findall(f'{{{W}}}style'):
        style_id = style_elem.get(f'{{{W}}}styleId')
        style_name = ''
        name_elem = style_elem.find(f'{{{W}}}name')
        if name_elem is not None:
            style_name = name_elem.get(f'{{{W}}}val', '')

        # 匹配 Pandoc 生成的脚注引用样式或标准 FootnoteReference 样式
        if style_id in (PANDOC_FOOTNOTE_REF_STYLE, 'FootnoteReference') or style_name == 'footnote reference':
            _ensure_font_size(style_elem, FOOTNOTE_FONT_SIZE)
            _ensure_superscript(style_elem)
            return

    # 如果没有找到，创建新样式
    _create_footnote_ref_style(styles_root)


def _ensure_font_size(style_elem, sz_val):
    """确保样式中包含字体大小设置"""
    rpr = style_elem.find(f'{{{W}}}rPr')
    if rpr is None:
        rpr = ET.SubElement(style_elem, f'{{{W}}}rPr')

    # 更新或添加 sz
    sz = rpr.find(f'{{{W}}}sz')
    if sz is None:
        sz = ET.SubElement(rpr, f'{{{W}}}sz')
    sz.set(f'{{{W}}}val', sz_val)

    # 更新或添加 szCs
    szcs = rpr.find(f'{{{W}}}szCs')
    if szcs is None:
        szcs = ET.SubElement(rpr, f'{{{W}}}szCs')
    szcs.set(f'{{{W}}}val', sz_val)


def _ensure_superscript(style_elem):
    """确保样式中包含上标设置"""
    rpr = style_elem.find(f'{{{W}}}rPr')
    if rpr is None:
        rpr = ET.SubElement(style_elem, f'{{{W}}}rPr')

    va = rpr.find(f'{{{W}}}vertAlign')
    if va is None:
        va = ET.SubElement(rpr, f'{{{W}}}vertAlign')
    va.set(f'{{{W}}}val', 'superscript')


def _ensure_paragraph_spacing(style_elem, after='0', line='240', line_rule='auto'):
    """确保样式中包含段落间距设置"""
    ppr = style_elem.find(f'{{{W}}}pPr')
    if ppr is None:
        ppr = ET.SubElement(style_elem, f'{{{W}}}pPr')

    spacing = ppr.find(f'{{{W}}}spacing')
    if spacing is None:
        spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
    spacing.set(f'{{{W}}}after', after)
    spacing.set(f'{{{W}}}line', line)
    spacing.set(f'{{{W}}}lineRule', line_rule)


def _create_footnote_text_style(styles_root):
    """创建脚注文字样式"""
    style = ET.SubElement(styles_root, f'{{{W}}}style')
    style.set(f'{{{W}}}type', 'paragraph')
    style.set(f'{{{W}}}styleId', 'FootnoteText')

    name = ET.SubElement(style, f'{{{W}}}name')
    name.set(f'{{{W}}}val', 'footnote text')
    based_on = ET.SubElement(style, f'{{{W}}}basedOn')
    based_on.set(f'{{{W}}}val', 'Normal')

    # 段落属性
    ppr = ET.SubElement(style, f'{{{W}}}pPr')
    spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
    spacing.set(f'{{{W}}}after', '0')
    spacing.set(f'{{{W}}}line', '240')
    spacing.set(f'{{{W}}}lineRule', 'auto')

    # 字体属性
    rpr = ET.SubElement(style, f'{{{W}}}rPr')
    sz = ET.SubElement(rpr, f'{{{W}}}sz')
    sz.set(f'{{{W}}}val', FOOTNOTE_FONT_SIZE)
    szcs = ET.SubElement(rpr, f'{{{W}}}szCs')
    szcs.set(f'{{{W}}}val', FOOTNOTE_FONT_SIZE)


def _create_footnote_ref_style(styles_root):
    """创建脚注引用样式"""
    style = ET.SubElement(styles_root, f'{{{W}}}style')
    style.set(f'{{{W}}}type', 'character')
    style.set(f'{{{W}}}styleId', 'FootnoteReference')

    name = ET.SubElement(style, f'{{{W}}}name')
    name.set(f'{{{W}}}val', 'footnote reference')

    # 字体属性
    rpr = ET.SubElement(style, f'{{{W}}}rPr')
    sz = ET.SubElement(rpr, f'{{{W}}}sz')
    sz.set(f'{{{W}}}val', FOOTNOTE_FONT_SIZE)
    szcs = ET.SubElement(rpr, f'{{{W}}}szCs')
    szcs.set(f'{{{W}}}val', FOOTNOTE_FONT_SIZE)
    va = ET.SubElement(rpr, f'{{{W}}}vertAlign')
    va.set(f'{{{W}}}val', 'superscript')
