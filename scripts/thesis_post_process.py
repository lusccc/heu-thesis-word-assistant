#!/usr/bin/env python3
"""
学位论文样式后处理脚本
功能：
1. 设置页边距
2. 格式化公式：使用eqArr结构，隐藏#号，右对齐编号
3. 自动识别章节，实现公式分章编号（如 1-1）
4. 更新交叉引用格式为"式（1-1）"，并清理冗余文本
"""
import sys
import os
import traceback
import re
from docx_processor.config import TEMP_DIR_NAME, DEFAULT_QMD_FILENAME
from docx_processor.utils import (
    extract_docx, 
    load_document_xml, 
    save_document_xml, 
    pack_docx, 
    cleanup_temp, 
    register_namespaces
)
from docx_processor.page_style import apply_page_settings
from docx_processor.equation_style import process_equations
from docx_processor.cross_reference import update_cross_references, strip_spaces_around_cross_refs
from docx_processor.header_footer import apply_header_footer
from docx_processor.bibliography import apply_bibliography_style, apply_phd_outcomes_style
from docx_processor.figure_style import process_figures
from docx_processor.table_style import process_tables
from docx_processor.toc_style import process_toc
from docx_processor.abstract_style import move_abstract_before_toc, bold_abstract_title, exclude_abstract_from_toc, set_keywords_heiti, apply_innovation_table_borders
from docx_processor.heading_style import remove_numbering_from_excluded_headings
from docx_processor.paragraph_style import process_paragraphs_with_math, strip_spaces_around_inline_math
from docx_processor.theorem_style import process_theorem_references
from docx_processor.equation_continuation import process_equation_continuation_paragraphs
from docx_processor.algorithm_style import process_algorithms
from docx_processor.symbol_style import process_circled_symbols
from docx_processor.landscape_table import process_landscape_tables
from docx_processor.footnote_style import process_footnotes
from docx_processor.config import read_qmd_yaml_bool
from docx_processor.title_pages import prepare_title_pages, merge_title_pages_content, compute_rid_map, merge_title_pages_files


def _extract_thesis_title_from_qmd(qmd_path: str):
    title = None
    try:
        with open(qmd_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except FileNotFoundError:
        return None

    # 仅解析YAML头（--- 到 ---）
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

        # 支持注释行：# title: "xxx"
        m = re.match(r'^#\s*title\s*:\s*(.+?)\s*$', stripped)
        if m:
            raw = m.group(1).strip()
            if (raw.startswith('"') and raw.endswith('"')) or (raw.startswith("'") and raw.endswith("'")):
                raw = raw[1:-1]
            title = raw.strip()
            break

        # 也支持非注释：title: "xxx"
        m = re.match(r'^title\s*:\s*(.+?)\s*$', stripped)
        if m:
            raw = m.group(1).strip()
            if (raw.startswith('"') and raw.endswith('"')) or (raw.startswith("'") and raw.endswith("'")):
                raw = raw[1:-1]
            title = raw.strip()
            break

    return title


def _resolve_qmd_path(docx_path: str, explicit_qmd_path: str | None = None) -> str:
    """解析与docx对应的QMD路径。

    优先级：
    1) 命令行显式传入的 qmd 路径
    2) 与 docx 同名的 qmd 文件
    3) docx 同目录下的默认 qmd 文件名
    4) 当前工作目录下的默认 qmd 文件名
    """
    if explicit_qmd_path:
        if os.path.exists(explicit_qmd_path):
            return explicit_qmd_path
        print(f"  警告：指定的QMD文件不存在，回退自动解析: {explicit_qmd_path}")

    docx_dir = os.path.dirname(docx_path)
    stem, _ = os.path.splitext(os.path.basename(docx_path))
    sibling_qmd = os.path.join(docx_dir, f"{stem}.qmd")
    if os.path.exists(sibling_qmd):
        return sibling_qmd

    qmd_in_docx_dir = os.path.join(docx_dir, DEFAULT_QMD_FILENAME)
    if os.path.exists(qmd_in_docx_dir):
        return qmd_in_docx_dir

    return os.path.abspath(DEFAULT_QMD_FILENAME)


def main():
    if len(sys.argv) < 2:
        print("用法：python thesis_post_process.py <docx文件路径> [qmd文件路径]")
        sys.exit(1)

    docx_path = sys.argv[1]
    explicit_qmd_path = sys.argv[2] if len(sys.argv) >= 3 else None

    try:
        # 注册命名空间
        register_namespaces()
        
        # 解压文件
        temp_path = extract_docx(docx_path, TEMP_DIR_NAME)
        
        # 加载XML
        tree, root = load_document_xml(temp_path)

        # 1. 设置页边距
        apply_page_settings(root)

        # 2. 处理目录样式（必须在处理页眉页脚之前）
        process_toc(root)

        # 2.5 将摘要和Abstract移到目录之前
        move_abstract_before_toc(root)

        # 2.6 排除摘要/Abstract出现在TOC中（设置outlineLvl=9）
        exclude_abstract_from_toc(root)

        # 2.7 Abstract标题加粗
        bold_abstract_title(root)

        # 2.8 关键词黑体字体
        set_keywords_heiti(root)

        # 3. 移除非正文章节的编号（结论、参考文献、致谢等）
        remove_numbering_from_excluded_headings(root)

        # 4. 处理公式并获取映射
        equation_map = process_equations(root)

        # 4. 处理图片样式和题注
        qmd_path = _resolve_qmd_path(docx_path, explicit_qmd_path)
        print(f"  使用QMD源文件: {qmd_path}")
        figure_map = process_figures(root, qmd_path)

        # 5. 处理表格样式和题注
        table_map = process_tables(root, qmd_path)

        # 5.1 创新成果自评表全边框（覆盖三线表样式）
        apply_innovation_table_borders(root)

        # 5.2 处理QCA符号：⊘ → 缩小的⊗（需在 thesis.qmd 中设置 qca-symbols: true）
        if read_qmd_yaml_bool(qmd_path, 'qca-symbols', default=False):
            process_circled_symbols(root)
        else:
            print("  QCA符号处理未启用（如需启用，请在 thesis.qmd YAML 头中设置 qca-symbols: true）")

        # 6. 更新交叉引用
        update_cross_references(root, equation_map)

        # 6.1 移除交叉引用前后的多余空格
        strip_spaces_around_cross_refs(root)

        # 7. 处理页眉页脚
        thesis_title = _extract_thesis_title_from_qmd(qmd_path)
        if thesis_title:
            print(f"  读取论文题目用于目录页眉: {thesis_title}")
        else:
            print("  未读取到论文题目，目录页眉将使用默认文本")
        apply_header_footer(temp_path, root, thesis_title=thesis_title)

        # 8. 处理定理环境交叉引用
        process_theorem_references(root)

        # 8.1 移除定理交叉引用前后的多余空格（定理引用在步骤8才创建为hyperlink，需在此之后去空格）
        strip_spaces_around_cross_refs(root)

        # 8.5 处理算法样式（中英文标题、三线表边框）
        process_algorithms(root, qmd_path)

        # 8.6 全局去除行内公式前后的多余空格
        strip_spaces_around_inline_math(root)

        # 9. 处理含公式段落的行间距
        process_paragraphs_with_math(root)

        # 9.1 处理公式后的续接段落（根据QMD源文件中是否有空行判断）
        process_equation_continuation_paragraphs(root, qmd_path)

        # 10. 处理参考文献样式
        apply_bibliography_style(root)

        # 11. 处理科研成果章节样式
        apply_phd_outcomes_style(root)

        # 11.5 横向页面表格（在页眉页脚处理之后，确保sectPr包含完整属性）
        process_landscape_tables(root, qmd_path)

        # 11.6 脚注处理：将尾注转换为脚注（①②③编号，每页重启，短横线分隔）
        process_footnotes(temp_path, root)

        # 12. 合并封面页（分两步：content在save前操作root，files在save后操作文件系统）
        docx_dir = os.path.dirname(os.path.abspath(docx_path))
        title_pages_path = os.path.join(docx_dir, "thesis_title_pages.docx")
        if not os.path.exists(title_pages_path):
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            title_pages_path = os.path.join(project_root, "thesis_title_pages.docx")
        
        tp_data = prepare_title_pages(title_pages_path)
        if tp_data is not None:
            print("正在合并封面页...")
            rid_map = compute_rid_map(temp_path, tp_data)
            merge_title_pages_content(root, tp_data, rid_map)

        # 保存XML
        save_document_xml(tree, temp_path)
        
        # 12b. 封面页文件操作（复制header/footer、更新rels和Content_Types）
        if tp_data is not None:
            merge_title_pages_files(temp_path, tp_data, rid_map)
        
        # 重新打包
        pack_docx(temp_path, docx_path)
        
        print("处理完成！")

    except Exception as e:
        print(f"处理过程中出错: {e}")
        traceback.print_exc()
        sys.exit(1)
    finally:
        # 清理临时文件
        cleanup_temp(TEMP_DIR_NAME)

if __name__ == "__main__":
    main()
