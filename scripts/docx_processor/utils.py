import xml.etree.ElementTree as ET
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import shutil

# 定义命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def register_namespaces():
    """注册命名空间以保持前缀"""
    for prefix, uri in NAMESPACES.items():
        ET.register_namespace(prefix, uri)
    # 额外注册其他常用命名空间（避免 ET 重写时丢失前缀、引发 Word 兼容告警）
    ET.register_namespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
    ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
    ET.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
    ET.register_namespace('w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex')
    ET.register_namespace('w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid')
    ET.register_namespace('w16', 'http://schemas.microsoft.com/office/word/2018/wordml')
    ET.register_namespace('w16cex', 'http://schemas.microsoft.com/office/word/2018/wordml/cex')
    ET.register_namespace('w16sdtdh', 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash')
    ET.register_namespace('w16sdtfl', 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock')
    ET.register_namespace('w16du', 'http://schemas.microsoft.com/office/word/2023/wordml/word16du')
    ET.register_namespace('o', 'urn:schemas-microsoft-com:office:office')
    ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')

# mc:Ignorable 前缀 → 命名空间 URI 的映射
# ET 重写 XML 时会丢弃未使用的命名空间声明，但 mc:Ignorable 属性以字符串
# 形式引用这些前缀。Word 发现 mc:Ignorable 引用了未声明的前缀时就会
# 报"无法读取的内容"。此映射用于在 ET.write 后补充缺失的 xmlns 声明。
_MC_PREFIX_TO_URI = {
    'r':          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w14':        'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15':        'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16':        'http://schemas.microsoft.com/office/word/2018/wordml',
    'w16cex':     'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'w16cid':     'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'w16du':      'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
    'w16sdtdh':   'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
    'w16sdtfl':   'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
    'w16se':      'http://schemas.microsoft.com/office/word/2015/wordml/symex',
}


def fix_mc_ignorable_namespaces(xml_path):
    """修复 ET.write 后丢失的 mc:Ignorable 所需命名空间声明。

    Python ElementTree 在写入 XML 时只声明实际使用的命名空间前缀，
    但 mc:Ignorable 属性以纯字符串引用前缀（如 "w14 w15 ..."），
    这些前缀如果没有 xmlns 声明，Word 会报"无法读取的内容"。

    本函数读取写入后的 XML，解析 mc:Ignorable，补充缺失的 xmlns 声明。
    """
    xml_path = Path(xml_path)
    if not xml_path.exists():
        return

    content = xml_path.read_text(encoding='utf-8')

    # 提取根元素开标签（跳过 <?xml?> 声明和注释）
    m = re.search(r'(<(?![\?!])[^>]+>)', content, re.S)
    if not m:
        return
    open_tag = m.group(1)

    # 提取 mc:Ignorable 的值
    ig_match = re.search(r'Ignorable="([^"]*)"', open_tag)
    if not ig_match:
        return

    prefixes_needed = ig_match.group(1).split()
    declared = set(re.findall(r'xmlns:(\w+)=', open_tag))

    missing_attrs = []
    for pfx in prefixes_needed:
        if pfx not in declared and pfx in _MC_PREFIX_TO_URI:
            missing_attrs.append(f'xmlns:{pfx}="{_MC_PREFIX_TO_URI[pfx]}"')

    if not missing_attrs:
        return

    # 在根元素开标签的 > 之前插入缺失的 xmlns 声明
    insertion = ' ' + ' '.join(missing_attrs)
    if open_tag.endswith('/>'):
        new_tag = open_tag[:-2] + insertion + '/>'
    else:
        new_tag = open_tag[:-1] + insertion + '>'

    content = content.replace(open_tag, new_tag, 1)
    xml_path.write_text(content, encoding='utf-8')


def extract_docx(docx_path, temp_dir):
    """解压 docx 文件"""
    docx_path = Path(docx_path)
    if not docx_path.exists():
        raise FileNotFoundError(f"文件 {docx_path} 不存在")
    
    temp_dir = Path(temp_dir)
    temp_dir.mkdir(exist_ok=True)
    
    with ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    return temp_dir

def pack_docx(temp_dir, original_path):
    """重新打包为 docx（直接覆盖原文件）"""
    original_path = Path(original_path)

    # 打包前批量修复 mc:Ignorable 引用的缺失命名空间声明
    for xml_file in temp_dir.rglob('*.xml'):
        fix_mc_ignorable_namespaces(xml_file)

    with ZipFile(original_path, 'w', ZIP_DEFLATED) as zip_ref:
        for file_path in temp_dir.rglob('*'):
            if file_path.is_file():
                arcname = file_path.relative_to(temp_dir)
                zip_ref.write(file_path, arcname)

    print(f"已覆盖输出文件：{original_path}")

def load_document_xml(temp_dir):
    """读取 document.xml"""
    doc_xml_path = temp_dir / "word" / "document.xml"
    tree = ET.parse(doc_xml_path)
    return tree, tree.getroot()

def save_document_xml(tree, temp_dir):
    """保存 document.xml"""
    doc_xml_path = temp_dir / "word" / "document.xml"
    tree.write(doc_xml_path, encoding='utf-8', xml_declaration=True)

def cleanup_temp(temp_dir):
    """清理临时目录"""
    if Path(temp_dir).exists():
        shutil.rmtree(temp_dir)


# === LaTeX 公式转 OMML 辅助函数 ===

LATEX_TO_UNICODE = {
    r'\alpha': '\u03b1', r'\beta': '\u03b2', r'\gamma': '\u03b3', r'\delta': '\u03b4',
    r'\epsilon': '\u03b5', r'\zeta': '\u03b6', r'\eta': '\u03b7', r'\theta': '\u03b8',
    r'\iota': '\u03b9', r'\kappa': '\u03ba', r'\lambda': '\u03bb', r'\mu': '\u03bc',
    r'\nu': '\u03bd', r'\xi': '\u03be', r'\pi': '\u03c0', r'\rho': '\u03c1',
    r'\sigma': '\u03c3', r'\tau': '\u03c4', r'\upsilon': '\u03c5', r'\phi': '\u03c6',
    r'\chi': '\u03c7', r'\psi': '\u03c8', r'\omega': '\u03c9',
    r'\Alpha': '\u0391', r'\Beta': '\u0392', r'\Gamma': '\u0393', r'\Delta': '\u0394',
    r'\Theta': '\u0398', r'\Lambda': '\u039b', r'\Pi': '\u03a0', r'\Sigma': '\u03a3',
    r'\Phi': '\u03a6', r'\Psi': '\u03a8', r'\Omega': '\u03a9',
    r'\times': '\u00d7', r'\pm': '\u00b1', r'\leq': '\u2264', r'\geq': '\u2265',
    r'\neq': '\u2260', r'\approx': '\u2248', r'\infty': '\u221e', r'\sim': '~',
}


def latex_to_omml_text(latex_str):
    """将简单的LaTeX公式字符串转换为用于OMML的Unicode文本"""
    result = latex_str.strip()
    for cmd, char in sorted(LATEX_TO_UNICODE.items(), key=lambda x: -len(x[0])):
        result = result.replace(cmd, char)
    return result


def create_omath_element(m_ns, w_ns, latex_content):
    """将LaTeX内容转换为OMML oMath元素"""
    omath = ET.Element(f'{{{m_ns}}}oMath')
    math_run = ET.SubElement(omath, f'{{{m_ns}}}r')
    mr_pr = ET.SubElement(math_run, f'{{{w_ns}}}rPr')
    mr_fonts = ET.SubElement(mr_pr, f'{{{w_ns}}}rFonts')
    mr_fonts.set(f'{{{w_ns}}}ascii', 'Cambria Math')
    mr_fonts.set(f'{{{w_ns}}}hAnsi', 'Cambria Math')
    mt = ET.SubElement(math_run, f'{{{m_ns}}}t')
    mt.text = latex_to_omml_text(latex_content)
    return omath


def append_text_with_math(para, w_ns, text):
    """向段落追加文本内容，自动将 $...$ 转为 OMML 数学元素
    
    自动去除公式前后与公式相邻的空格，避免在Word中产生多余间距。
    
    Args:
        para: 段落元素
        w_ns: Word命名空间URI
        text: 可含 $...$ 公式的文本
    """
    m_ns = NAMESPACES['m']
    parts = re.split(r'(\$[^$]+\$)', text)
    
    # 去除公式前后相邻文本段的空格
    for i in range(len(parts)):
        if not parts[i]:
            continue
        is_math = parts[i].startswith('$') and parts[i].endswith('$')
        if is_math:
            continue
        # 如果前一个非空 part 是公式，去除当前文本段开头空格
        for j in range(i - 1, -1, -1):
            if parts[j]:
                if parts[j].startswith('$') and parts[j].endswith('$'):
                    parts[i] = parts[i].lstrip(' ')
                break
        # 如果后一个非空 part 是公式，去除当前文本段末尾空格
        for j in range(i + 1, len(parts)):
            if parts[j]:
                if parts[j].startswith('$') and parts[j].endswith('$'):
                    parts[i] = parts[i].rstrip(' ')
                break
    
    for part in parts:
        if not part:
            continue
        if part.startswith('$') and part.endswith('$'):
            latex_content = part[1:-1]
            omath = create_omath_element(m_ns, w_ns, latex_content)
            para.append(omath)
        else:
            run = ET.Element(f'{{{w_ns}}}r')
            t = ET.SubElement(run, f'{{{w_ns}}}t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = part
            para.append(run)
