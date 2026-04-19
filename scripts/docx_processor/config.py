"""
论文处理配置文件
集中管理所有可配置的常量
"""

# 学校和论文信息
UNIVERSITY_NAME = "哈尔滨工程大学"
THESIS_TYPE = "博士学位论文"

# 页边距配置 (单位: mm)
# 1 mm = 56.7 twips
MM_TO_TWIPS = 56.7

PAGE_MARGINS = {
    "top": 28,      # 上边距 28mm
    "bottom": 28,   # 下边距 28mm
    "left": 25,     # 左边距 25mm
    "right": 25,    # 右边距 25mm
    "header": 20,   # 页眉距离 20mm
    "footer": 20,   # 页脚距离 20mm
    "gutter": 0,    # 装订线 0mm
}

# 字体配置
MATH_FONT = "Latin Modern Math"

# 排除的章节标题（不计入章节编号）
EXCLUDE_CHAPTER_TITLES = [
    '摘要', 'Abstract', 
    '目录', 'Table of contents',
    '结论', 'Conclusion',
    '致谢', 'Acknowledgement',
    '参考文献', 'References',
    '读博士学位期间发表的论文和取得的科研成果',
    '附录',
    '博士学位论文创新成果自评表'
]

# 样式ID映射（来自模板）
STYLES = {
    "heading1": "1",           # 一级标题
    "heading1_alt": "Heading1", # 一级标题备选
    "bibliography": "af7",      # 参考文献
    "figure_caption": "af3",    # 图片题注
    "figure": "af6",            # 图（包含图片的段落）
    "header": "ae",             # 页眉
    "footer": "af0",            # 页脚
    "equation": "equation-style", # 公式
    "image_caption": "ImageCaption", # 图片标题（Quarto生成）
    "caption": "af3",           # 题注（表格标题段落样式）
    "table_content": "af5",     # 表（表格内容段落样式）
    "three_line_table": "aff0", # 三线表（表格设计样式）
}

# 三线表边框配置（用于自定义边框表格的基础三线，通过 tcBorders 显式写入）
# sz 单位为 1/8 磅 (eighth of a point)
THREE_LINE_BORDERS = {
    'top':            {'val': 'single', 'sz': '12', 'color': '000000', 'space': '0'},  # 顶线 1.5pt
    'header_bottom':  {'val': 'single', 'sz': '6',  'color': '000000', 'space': '0'},  # 表头底线 0.75pt
    'bottom':         {'val': 'single', 'sz': '12', 'color': '000000', 'space': '0'},  # 底线 1.5pt
}

# 参考文献标题关键词
BIBLIOGRAPHY_TITLE = "参考文献"

# 临时目录名
TEMP_DIR_NAME = "temp_docx_processing"

# 公式段落制表位（用于 m:eqArr + # 分隔符的对齐）
# A4 可用宽度 = 页面宽度 - 左边距 - 右边距
_PAGE_WIDTH_TWIPS = 11906  # A4 宽度
_AVAILABLE_WIDTH = _PAGE_WIDTH_TWIPS - int(PAGE_MARGINS['left'] * MM_TO_TWIPS) - int(PAGE_MARGINS['right'] * MM_TO_TWIPS)
EQUATION_TAB_CENTER = str(_AVAILABLE_WIDTH // 2)  # 居中制表位（公式位置）
EQUATION_TAB_RIGHT = str(_AVAILABLE_WIDTH)         # 右对齐制表位（编号位置）

# 默认QMD文件名
DEFAULT_QMD_FILENAME = "thesis.qmd"


def get_margin_twips(margin_name):
    """获取边距的twips值"""
    mm = PAGE_MARGINS.get(margin_name, 0)
    return str(int(mm * MM_TO_TWIPS))


def read_qmd_yaml_bool(qmd_path, key, default=False):
    """从 QMD 文件的 YAML 头中读取布尔配置项

    支持格式：
      key: true/false        ← 生效
      # key: true/false      ← 被注释，视为未设置，返回 default

    注释行仍会被解析，便于用户快速切换：
      取消注释即可启用，加注释即可禁用。

    Args:
        qmd_path: QMD 文件路径
        key: YAML 键名（如 'qca-symbols'）
        default: 未找到时的默认值

    Returns:
        bool
    """
    import os
    import re
    if not qmd_path or not os.path.exists(qmd_path):
        return default
    try:
        with open(qmd_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except (FileNotFoundError, OSError):
        return default

    in_yaml = False
    for line in lines:
        stripped = line.strip()
        if stripped == '---':
            if not in_yaml:
                in_yaml = True
                continue
            break
        if not in_yaml:
            continue

        # 仅匹配非注释行：key: true/false
        m = re.match(
            r'^' + re.escape(key) + r'\s*:\s*(true|false)\s*$',
            stripped,
        )
        if m:
            return m.group(1).lower() == 'true'

        # 也匹配注释行，用于检测用户意图但不生效
        # （仅做占位，不返回值，让函数继续搜索或返回 default）

    return default
