# HEU学位论文排版助手

> 哈尔滨工程大学博士/硕士学位论文 DOCX 自动排版工具
>
> 本工具按照哈尔滨工程大学官方文件《研究生学位论文撰写规范》制作，自动处理页边距、页眉页脚、公式编号、三线表、脚注等格式要求。哈尔滨工程大学本科生院和研究生院提供了规范和以及本科生word模板，此模板仅为规范的参考实现，不保证格式审查老师不提意见。任何由于使用本模板而引起的论文格式审查问题均与本模板作者无关
>
> 友情链接：[HeuThesis](https://github.com/Li-Wenhui/HEUThesis)

---

## 快速开始

### 0. 环境配置（首次使用）

**安装 Quarto**（含 Pandoc）：从 [quarto.org](https://quarto.org/docs/get-started/) 下载安装，安装后确认：

```bash
quarto --version   # 应输出 ≥ 1.4
```

**安装 Python**：需 3.10+，无需任何第三方包。确认：

```bash
python3 --version  # Linux/WSL
python --version   # Windows
```

**安装 Microsoft Word**：2016+，用于最终审阅和 VBA 宏。

**配置编辑器**：推荐使用以下编辑器，均基于 VS Code 架构，需安装 **Quarto 扩展**以获得 `.qmd` 语法高亮、预览和智能补全：

| 编辑器             | Quarto 扩展安装方式                               |
| ------------------ | ------------------------------------------------- |
| **VS Code**  | 扩展商店搜索 `Quarto` 安装                      |
| **Windsurf** | 扩展商店搜索 `Quarto` 安装（兼容 VS Code 扩展） |
| **Cursor**   | 扩展商店搜索 `Quarto` 安装（兼容 VS Code 扩展） |
| **Trae**     | 扩展商店搜索 `Quarto` 安装（兼容 VS Code 扩展） |
| **Qoder**    | 扩展商店搜索 `Quarto` 安装（兼容 VS Code 扩展） |

> 提示：这些编辑器均内置终端，可直接运行编译脚本。Quarto 扩展提供 `.qmd` 文件的语法高亮、实时预览（`Ctrl+Shift+K`）和 YAML 智能提示。

### 1. 写论文

编辑论文封面 ` thesis_title_pages.docx`；用编辑器打开  `thesis.qmd`，用 Markdown 语法撰写论文内容。

### 2. 编译

打开终端/命令行输入命令：

```bash
# Linux / WSL
bash compile_thesis_linux.sh

# Windows
compile_thesis_windows.bat
```

编译完成后，`output/thesis.docx` 即为排版后的论文。

> **快捷编译：** 项目已配置 VS Code Build Task，在 VS Code / Windsurf / Cursor / Trae / Qoder 中按 **`Ctrl+Shift+B`** 即可一键编译（Linux/WSL 默认运行 `Compile: Thesis (Linux/WSL)`，Windows 用户需从命令面板 `Ctrl+Shift+P` → `Run Build Task` 选择 `Compile: Thesis (Windows)`）。

### 3. Word 定稿

在 Word 中打开 `output/thesis.docx`，按顺序运行 VBA 宏：

1. `AddContinuedTableLabels`（续表标签）
2. `InsertBlankPages`（空白页插入）

> 部分精确排版（如图片位置微调、行高调整等）仍需在 Word 中手动完成，详见 [第10节](#10-精确排版与手动调整)。

### 4. 导出 PDF

> ⚠️ **重要提示：导出 PDF 必须使用 Adobe Acrobat Pro DC，切勿使用 Word 自带的"另存为 PDF"功能！**
>
> 1. 安装 [Adobe Acrobat Pro DC](https://www.gndown.com/11177.html)
> 2. 在 Word 中点击 **文件 → 另存为 Adobe PDF** 导出
>
> Word 自带的 PDF 导出会导致公式格式出错，Adobe PDF 能完整保留 DOCX 的排版效果。

### 5. 引号转换（可选）

中文论文要求使用弯引号（`“”`），但键盘直接输入的是直引号（`"`）。项目提供了 `smart_quotes_cn.py` 脚本，可自动将 `thesis.qmd` 中的直双引号转换为正确的中文弯双引号，同时保护代码块、公式、HTML 标签等不被误改：

```bash
python3 smart_quotes_cn.py thesis.qmd
```

脚本会自动备份原文件为 `thesis.qmd.bak`，然后原地替换。默认处理 `thesis.qmd`，也可指定其他文件。

---

## 目录

- [快速开始](#快速开始)
  - [0. 环境配置](#0-环境配置首次使用)
  - [1. 写论文](#1-写论文)
  - [2. 编译](#2-编译)
  - [3. Word 定稿](#3-word-定稿)
  - [4. 导出 PDF](#4-导出-pdf)
  - [5. 引号转换](#5-引号转换可选)
- [1. 项目简介](#1-项目简介)
- [2. 原理与目的](#2-原理与目的)
- [3. 环境与安装](#3-环境与安装)
- [4. 快速开始（详细版）](#4-快速开始详细版)
- [5. 项目结构](#5-项目结构)
- [6. 编译脚本](#6-编译脚本)
- [7. thesis.qmd 编写指南](#7-thesisqmd-编写指南)
  - [7.1 YAML 头配置](#71-yaml-头配置)
  - [7.2 章节结构与标记](#72-章节结构与标记)
  - [7.3 公式](#73-公式)
  - [7.4 图片](#74-图片)
  - [7.5 表格](#75-表格)
  - [7.6 定理环境](#76-定理环境)
  - [7.7 脚注](#77-脚注)
  - [7.8 参考文献](#78-参考文献)
  - [7.9 QCA 符号](#79-qca-符号)
- [8. 封面页定制](#8-封面页定制)
- [9. VBA 宏使用指南](#9-vba-宏使用指南)
  - [9.1 运行方法](#91-运行方法)
  - [9.2 执行顺序（关键）](#92-执行顺序关键)
  - [9.3 各宏详解](#93-各宏详解)
- [10. 精确排版与手动调整](#10-精确排版与手动调整)
- [11. 后处理流水线详解](#11-后处理流水线详解)
- [12. Lua 过滤器说明](#12-lua-过滤器说明)
- [13. 模板文件](#13-模板文件)
- [14. 常见问题与排错](#14-常见问题与排错)
- [15. 扩展指南](#15-扩展指南)

---

## 1. 项目简介

**HEU学位论文排版助手** 是一套面向哈尔滨工程大学博士学位论文的自动化排版工具。用户只需在 `thesis.qmd` 中用 Markdown 语法撰写论文内容，运行编译脚本即可生成符合学校格式要求的 DOCX 文件，无需手动设置页边距、页眉页脚、公式编号、三线表样式等繁琐格式。

**核心特色：**

- 自动设置页边距（上下28mm、左右25mm）
- 公式分章编号（如 (2-1)）+ 右对齐 + 交叉引用
- 三线表样式自动应用 + 英文表标题 + 表注
- 页眉页脚自动配置（奇偶页不同、章节标题页眉）
- 脚注自动转为①②③编号 + 每页重启
- 定理环境交叉引用（手动 bookmark 语法）
- QCA 符号自动处理（核心/边缘条件大小区分）
- 封面页自动合并
- 横向表格自动旋转
- 创新成果自评表全边框

## 2. 原理与目的

### 目的

博士学位论文格式要求严格，手动在 Word 中调整格式耗时且易出错。本项目旨在：

1. **让作者专注于内容**：用 Markdown 写论文，格式由工具自动处理
2. **一键编译出DOCX**：运行一个脚本即可生成排版后的论文
3. **可重复构建**：修改内容后重新编译，格式自动恢复，不会因手动操作丢失

### 原理

本项目采用 **Quarto 渲染 + Python 后处理 + VBA 定稿** 三阶段流水线：

```
thesis.qmd ──Quarto──→ 原始DOCX ──Python后处理──→ 排版DOCX ──VBA宏──→ 定稿DOCX
```

**第一阶段：Quarto 渲染**

Quarto 将 Markdown（`.qmd`）通过 Pandoc 转换为 DOCX。此阶段处理基本的内容转换（标题、段落、公式、图片、表格、参考文献等），同时通过 Lua 过滤器完成：

- 公式分章编号与交叉引用链接
- 一级标题前自动插入分节符
- 横向表格页面旋转

> 定理环境（定义/引理/定理/证明）目前采用手动 bookmark 语法编写，后处理脚本负责交叉引用链接。如需 Div 语法自动编号，可将 `filters/theorem_env.lua` 加入 YAML filters 列表。

**第二阶段：Python 后处理**

Python 脚本解压 DOCX（本质是 ZIP），直接操作内部 XML 文件，完成 Quarto/Pandoc 无法实现的精细排版：

- 页边距、页眉页脚、奇偶页设置
- 三线表样式、自定义边框
- 脚注①②③编号、每页重启
- QCA 符号字体统一与大小区分
- 公式编号右对齐、续接段落缩进
- 封面页内容合并
- OOXML 命名空间修复

**第三阶段：VBA 定稿**

在 Word 中运行 VBA 宏，处理 Python 无法完成的需要 Word 排版引擎参与的操作：

- 跨页表格分割 + 续表标签
- 章节标题奇数页起始（空白页插入）
- 图片/表格宽度信息提取（辅助工具）

> **为什么需要三个阶段？** Quarto 负责内容转换，Python 负责底层 XML 精细控制，VBA 负责需要 Word 排版引擎的计算（如页码判断、分页位置检测）。三者各司其职，无法互相替代。

## 3. 环境与安装

### 必需软件

| 软件                     | 版本要求 | 说明                                                       |
| ------------------------ | -------- | ---------------------------------------------------------- |
| **Quarto**         | ≥ 1.4   | [下载地址](https://quarto.org/docs/get-started/)，包含 Pandoc |
| **Python**         | ≥ 3.10  | 无第三方依赖，纯标准库                                     |
| **Microsoft Word** | ≥ 2016  | 用于最终审阅和 VBA 宏                                      |

### 推荐编辑器

本项目的核心源文件 `thesis.qmd` 是 Markdown 格式，推荐使用以下编辑器，可获得智能补全、实时预览和终端集成：

- **VS Code** + Quarto 扩展
- **Windsurf**
- **Cursor**
- **Qoder**
- **Trae**

这些编辑器均内置终端，可直接运行编译脚本，并支持 YAML/Markdown 语法高亮。

### 安装步骤

1. 安装 Quarto（含 Pandoc）
2. 确认 Python 3.10+ 已安装：`python3 --version`（Linux）或 `python --version`（Windows）
3. 克隆本项目
4. 无需安装 Python 依赖包——本项目仅使用标准库

## 4. 快速开始（详细版）

### 4.1 编辑论文内容

打开 `thesis.qmd`，修改 YAML 头中的论文信息：

```yaml
---
# title: "你的论文题目"
# author: "你的姓名"
# date: "2026年1月"
---
```

> 注意：`title`/`author`/`date` 前的 `#` 是注释，这是故意的——这些信息由封面页 `thesis_title_pages.docx` 提供，YAML 中的值仅用于页眉等辅助位置。

然后编写论文正文。

### 4.2 编译

**Linux / WSL：**

```bash
bash compile_thesis_linux.sh
```

**Windows：**

```cmd
compile_thesis_windows.bat
```

编译完成后，`output/thesis.docx` 即为排版后的论文。

### 4.3 VBA 定稿

在 Word 中打开 `output/thesis.docx`，按顺序运行 VBA 宏（详见 [第9节](#9-vba-宏使用指南)）。

## 5. 项目结构

```
heu论文助手/
├── thesis.qmd                    # 论文源文件（Markdown + YAML）
├── thesis_ref.bib                # 参考文献数据库
├── thesis_title_pages.docx       # 封面页源文件
├── smart_quotes_cn.py            # 直引号→弯引号转换工具
├── compile_thesis_linux.sh       # Linux/WSL 编译脚本
├── compile_thesis_windows.bat    # Windows 编译脚本
├── output/
│   └── thesis.docx               # 编译输出
├── docx_template/
│   └── heu_thesis_style.docx     # Quarto 参考模板（定义样式ID）
├── filters/
│   ├── eq_number.lua             # 公式分章编号
│   ├── next_page.lua             # 一级标题分节符
│   ├── landscape.lua             # 横向表格（Lua方式，本项目未启用）
│   ├── page_break.lua            # 分页符
│   └── theorem_env.lua           # 定理环境（Lua方式，本项目未启用）
├── scripts/
│   ├── thesis_post_process.py    # 后处理主入口
│   └── docx_processor/           # 后处理模块
│       ├── config.py             # 集中配置
│       ├── utils.py              # 解压/打包/命名空间
│       ├── page_style.py         # 页边距
│       ├── equation_style.py     # 公式排版
│       ├── figure_style.py       # 图片样式
│       ├── table_style.py        # 表格样式
│       ├── header_footer.py      # 页眉页脚
│       ├── footnote_style.py     # 脚注处理
│       ├── symbol_style.py       # QCA 符号
│       ├── title_pages.py        # 封面合并
│       ├── toc_content.py        # TOC 内容提取
│       └── __init__.py
├── vba/
│   ├── ContinuedTable.vba        # 续表标签
│   ├── InsertBlankPages.vba      # 空白页插入
│   ├── GetImageWidth.vba         # 图片宽度提取
│   ├── GetTableColWidth.vba      # 表格列宽查看
│   └── FieldCodeToText.vba       # 域代码转文本（调试用）
└── gbt7714-2015-numeric-bilingual.csl  # GB/T 7714 引用格式
```

## 6. 编译脚本

### compile_thesis_linux.sh

适用于 Linux 和 WSL 环境。脚本流程：

1. 创建 `output/` 目录
2. `quarto render thesis.qmd --to docx` 渲染
3. 移动输出文件到 `output/thesis.docx`
4. `python3 scripts/thesis_post_process.py` 后处理
5. （WSL 环境）通过 PowerShell 自动用 Word 打开

### compile_thesis_windows.bat

适用于原生 Windows 环境。脚本流程：

1. 创建 `output/` 目录
2. `quarto render thesis.qmd --to docx` 渲染
3. 移动输出文件到 `output/thesis.docx`
4. `python scripts\thesis_post_process.py` 后处理
5. `start` 命令用 Word 打开

> 注意：Windows 脚本使用 `python`（非 `python3`），路径分隔符为 `\`。

### 手动编译

如需分步执行：

```bash
# 1. Quarto 渲染
quarto render thesis.qmd --to docx --output thesis.docx

# 2. 移动到 output 目录
mv thesis.docx output/thesis.docx

# 3. 后处理
python3 scripts/thesis_post_process.py output/thesis.docx thesis.qmd
```

## 7. thesis.qmd 编写指南

### 7.1 YAML 头配置

```yaml
---
# title: "论文题目"          # 注释行，值用于页眉等辅助位置
# author: "作者姓名"          # 注释行
# date: "2026年1月"           # 注释行
format:
  docx:
    reference-doc: docx_template/heu_thesis_style.docx  # 模板文件
    toc: true                          # 生成目录
    number-sections: false              # 章节编号由后处理控制
bibliography: thesis_ref.bib           # 参考文献数据库
csl: gbt7714-2015-numeric-bilingual.csl  # GB/T 7714 格式
link-citations: false                  # 不链接引用编号
qca-symbols: true                      # 启用 QCA 符号处理（不用则删除或设 false）
filters:
  - filters/eq_number.lua              # 公式编号
  - filters/next_page.lua              # 分节符
  - filters/landscape.lua              # 横向表格
---
```

**配置项说明：**

| 配置项                          | 必填 | 说明                                            |
| ------------------------------- | ---- | ----------------------------------------------- |
| `format.docx.reference-doc`   | 是   | Word 模板，定义样式ID映射                       |
| `format.docx.toc`             | 是   | 必须为 `true`，后处理会修改目录样式           |
| `format.docx.number-sections` | 是   | 必须为 `false`，编号由 Lua 过滤器和后处理控制 |
| `bibliography`                | 是   | BibTeX 文件路径                                 |
| `csl`                         | 是   | 引文格式文件                                    |
| `qca-symbols`                 | 否   | `true` 启用 QCA 符号处理，默认 `false`      |
| `filters`                     | 是   | 三个 Lua 过滤器，顺序不可更改                   |

### 7.2 章节结构与标记

论文必须按以下顺序组织章节：

```markdown
# 摘    要 {.unnumbered .unlisted}
（摘要内容）

# Abstract {.unnumbered .unlisted}
（Abstract内容）

# 博士学位论文创新成果自评表 {.unnumbered .unlisted}
（自评表内容）

（目录由 Quarto 自动生成，无需手动编写）

# 绪论 {#sec-intro}
（正文第一章）

# 理论基础与研究方法 {#sec-theory}
（正文第二章）

...（更多正文章节）

# 结论与展望 {#sec-conclusion}
（结论）

# 附录：XXX {#sec-appendix-xxx .unnumbered}
（附录，如有）

# 参考文献

[^footnote-id]: 脚注内容

::: {#refs}
:::

# 读博士学位期间发表的论文和取得的科研成果 {#sec-phd-outcomes .unnumbered}
（科研成果）

# 致    谢 {#sec-ack .unnumbered}
（致谢内容）
```

**关键说明：**

- `.unnumbered .unlisted`：不编号、不出现在目录中。用于摘要、Abstract、致谢等非正文章节
- 正文章节（一级标题）不加 `.unnumbered`，会自动分章编号
- `# 参考文献` 和 `::: {#refs} :::` 必须成对出现
- "致    谢"中间有4个空格，这是学校格式要求
- 后处理会自动识别并禁用以下章节的编号：摘要、Abstract、目录、结论、参考文献、致谢、科研成果、附录、创新成果自评表

### 7.3 公式

**行内公式：**

```markdown
黄金价格序列为 $\{p_t\}_{t=1}^{T}$，预测窗口为 $h$。
```

**行间公式（带编号）：**

```markdown
$$ \hat{p}_{T+h} = f(p_{T-w+1}, \ldots, p_T) $$ {#eq-lstm-forget}
```

- 编号格式自动为 `(章号-序号)`，如 `(2-1)`
- `{#eq-xxx}` 为公式标签，用于交叉引用

**交叉引用：**

```markdown
由式 @eq-lstm-forget 可知...
```

渲染结果：由式（2-1）可知...

**续接段落：**

公式后如果没有空行，下一段文字被视为公式的续接说明，后处理会自动设置首行缩进：

```markdown
$$ E = mc^2 $$ {#eq-energy}
其中 $E$ 为能量，$m$ 为质量，$c$ 为光速。
```

**公式后有空行则为独立段落：**

```markdown
$$ E = mc^2 $$ {#eq-energy}

下面是独立段落，不缩进。
```

### 7.4 图片

**基本语法：**

```markdown
![研究框架](images/framework.png){#fig-framework}
```

**英文标题：**

在图片后添加 HTML 注释：

```markdown
![研究框架](images/framework.png){#fig-framework}

<!-- fig-cap-en: Research framework -->
```

**宽度控制：**

```markdown
![研究框架](images/framework.png){#fig-framework width="80%"}
```

**交叉引用：**

```markdown
如图 @fig-framework 所示...
```

渲染结果：如图2.1所示...

> 图片宽度百分比可通过 VBA 宏 `提取图片宽度百分比` 从 Word 中提取后回填到 QMD（见 [9.3节](#93-各宏详解)）。

### 7.5 表格

#### Pipe Table（简单表格）

适用于不需要自定义边框的简单表格：

```markdown
| 条件 | 路径1 | 路径2 |
|:-----|:-----:|:-----:|
| 经济增长 | ⬤ | ⊗ |
| 通胀预期 | ⊗ | ⬤ |

: QCA符号示例 {#tbl-qca-example}
```

#### HTML Table（自定义边框/列对齐）

需要自定义单元格边框、精确列宽或行合并时，使用 HTML 表格：

```markdown
::: {#tbl-data-sources}

```{=html}
<table>
  <caption>黄金市场预测的多源数据分类</caption>
  <colgroup>
    <col style="width: 12%;">
    <col style="width: 15%;">
  </colgroup>
  <thead>
    <tr>
      <th><span data-qmd="数据来源"></span></th>
      <th><span data-qmd="数据类别"></span></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="4"><span data-qmd="市场数据"></span></td>
      <td><span data-qmd="黄金价格"></span></td>
    </tr>
    <tr style="border-bottom: 0.75pt solid black">
      <td><span data-qmd="汇率数据"></span></td>
    </tr>
  </tbody>
</table>
```

:::

```

**关键语法说明：**

- `<caption>` 标签提供表格中文标题（替代 Pipe Table 的 `: 标题` 语法）
- `<span data-qmd="文本">` 确保中文内容被 Pandoc 正确处理
- `<colgroup>` + `<col style="width: X%;">` 设置列宽百分比
- `rowspan="N"` 实现行合并
- `style="border-bottom: ..."` 自定义单元格边框
- 后处理会解析 HTML 中的 `style` 属性，在 DOCX 中叠加自定义边框

#### 英文表标题

在表格标题行后添加 HTML 注释：

```markdown
: 黄金市场预测的多源数据分类 {#tbl-data-sources}

<!-- tbl-cap-en: Classification of multi-source data for gold market prediction -->
```

HTML 表格中使用 `<caption>` 提供中文标题时，英文标题同样用 `<!-- tbl-cap-en: ... -->` 注释添加在 Div 结束标记 `:::` 之后。

#### 表注

```markdown
: QCA符号示例 {#tbl-qca-example}

<!-- tbl-note: ⬤核心条件存在，⊗核心条件缺失。一致性≥0.80为可接受阈值。 -->
```

表注会显示在表格下方，格式为宋体小五号。

#### 横向表格

当表格列数较多、纵向页面放不下时，添加横向标记：

```markdown
<!-- landscape: 黄金价格预测 -->

::: {#tbl-ch3-main}
（表格内容）
:::
```

`landscape:` 后的关键词用于后处理识别该表格。编译后该表格会自动旋转为横向页面，前后页面保持纵向。

#### 交叉引用

```markdown
由表 @tbl-data-sources 可知...
```

渲染结果：由表2.1可知...

### 7.6 定理环境

本项目的 `thesis.qmd` 采用**手动 bookmark 语法**编写数学环境，即用加粗文本 + 行内链接标签的方式声明定义、引理、定理等，后处理脚本负责将引用转为超链接。

> **备选方案：** 如需自动分章编号，可将 `filters/theorem_env.lua` 加入 YAML 的 `filters` 列表，改用 Div 语法（见 [第12节](#12-lua-过滤器说明)）。

**定义：**

```markdown
[**定义 2.1（黄金价格预测）**]{#def-gold-price} 设黄金价格时间序列为 $\{p_t\}_{t=1}^{T}$，预测目标是学习映射函数 $f$...
```

**引理：**

```markdown
[**引理 3.1（注意力增强的信息量）**]{#lem-attn-info} 设LSTM编码器输出为 $\boldsymbol{H}$，多头注意力输出为 $\boldsymbol{A}$。则：
```

**定理（带自定义名称）：**

```markdown
[**定理 3.1（多维输入的预测增益）**]{#thm-multi-input} 设仅使用价格特征的预测模型信息量为 $I(\boldsymbol{X}_p; Y)$...
```

渲染结果：**定理 3.1（多维输入的预测增益）** 设仅使用...

**证明：**

证明不使用 bookmark 语法，直接用加粗文本 + 手动 QED 符号：

```markdown
**证明** 由于 $\boldsymbol{A} = g(\boldsymbol{H})$ 为确定性函数...即得所需结论。 □
```

**语法要点：**

- `[**类型 编号（名称）**]{#标签}` — 加粗文本 + bookmark 锚点
- 编号需手动编写（如 `2.1`、`3.1`），格式为 `章号.序号`
- 证明不编号，以 `□` 结尾

**交叉引用：**

```markdown
由 @thm-multi-input 可知...
```

渲染结果：由定理3.1可知...

**支持的环境类型：**

| 类型 | 编号格式 | 示例                                |
| ---- | -------- | ----------------------------------- |
| 定义 | 分章编号 | `[**定义 2.1（...）**]{#def-xxx}` |
| 引理 | 分章编号 | `[**引理 3.1（...）**]{#lem-xxx}` |
| 定理 | 分章编号 | `[**定理 3.1（...）**]{#thm-xxx}` |
| 推论 | 分章编号 | `[**推论 3.1（...）**]{#cor-xxx}` |
| 证明 | 不编号   | `**证明** ... □`                 |

### 7.7 脚注

使用标准 Markdown 脚注语法：

```markdown
正文中的脚注标记[^nixon-shock]。

[^nixon-shock]: 1971年8月15日，美国总统尼克松宣布停止美元与黄金的兑换。
```

后处理自动将脚注转为：

- ①②③带圈数字编号
- 每页重启编号
- 小五号字体
- 短横线分隔符

### 7.8 参考文献

**引用格式：**

```markdown
研究表明[@smith1997traffic; @williams2003modeling]...
```

**参考文献数据库：** 编辑 `thesis_ref.bib`，添加 BibTeX 条目：

```bibtex
@article{smith1997traffic,
  title   = {Traffic flow dynamics and the forecasting of traffic volumes},
  author  = {Smith, Brian L and Demetsky, Michael J},
  journal = {Transportation Research Record},
  volume  = {1603},
  number  = {1},
  pages   = {73--80},
  year    = {1997}
}
```

**引用格式标准：** 使用 `gbt7714-2015-numeric-bilingual.csl`，符合 GB/T 7714-2015 数字编号式双语格式。

**参考文献章节：**

```markdown
# 参考文献

[^nixon-shock]: 1971年8月15日...

::: {#refs}
:::
```

`::: {#refs} :::` 是 Pandoc 生成参考文献列表的占位符。脚注定义可放在参考文献标题下方。

### 7.9 QCA 符号

当论文包含 QCA（定性比较分析）表格时，在 YAML 头中启用：

```yaml
qca-symbols: true
```

在表格中使用四种符号：

| 符号 | Unicode | 含义         | 后处理效果                |
| ---- | ------- | ------------ | ------------------------- |
| ⬤   | U+2B24  | 核心条件存在 | 统一 Segoe UI Symbol 字体 |
| ⊗   | U+2297  | 核心条件缺失 | 放大115% + 统一字体       |
| ●   | U+25CF  | 边缘条件存在 | 缩小至8pt + 统一字体      |
| ⊘   | U+2298  | 边缘条件缺失 | 替换为缩小⊗ + 统一字体   |

**示例：**

```markdown
| 条件 | 路径1 | 路径2 | 路径3 |
|:-----|:-----:|:-----:|:-----:|
| 经济增长 | ⬤ | ⊗ | ● |
| 通胀预期 | ⊗ | ⬤ | ⊘ |

: QCA符号示例 {#tbl-qca-example}

<!-- tbl-cap-en: Condition combinations for gold price increase (QCA symbol example) -->

<!-- tbl-note: ⬤核心条件存在，⊗核心条件缺失，●边缘条件存在，⊘边缘条件缺失。 -->
```

核心/边缘条件通过**符号大小**区分：核心条件为正常字号（10.5pt），边缘条件缩小至8pt。⊗字形偏小，额外放大115%以匹配同级符号。

## 8. 封面页定制

封面页内容来自 `thesis_title_pages.docx`，在编译最后阶段自动合并到论文正文之前。

**编辑方法：**

1. 用 Word 打开 `thesis_title_pages.docx`
2. 修改论文题目、作者、导师、学科专业、日期等信息
3. 保存

> 注意：不要修改封面页中的分节符和页面设置，这些由后处理脚本依赖。

封面页合并过程会自动：

- 复制封面页中的 header/footer 文件
- 分配新的关系ID（rId）
- 更新 `[Content_Types].xml` 和 `word/_rels/document.xml.rels`

## 9. VBA 宏使用指南

VBA 宏用于处理需要 Word 排版引擎参与的操作，是定稿的必要步骤。

### 9.1 运行方法

1. 在 Word 中打开编译后的 `output/thesis.docx`
2. 按 **Alt+F11** 打开 VBA 编辑器
3. 点击 **插入 → 模块**
4. 粘贴 VBA 代码
5. 按 **F5** 运行

### 9.2 执行顺序（关键）

⚠️ **VBA 宏必须按以下顺序执行，顺序错误会导致排版错位：**

```
1. AddContinuedTableLabels   （续表标签）
2. InsertBlankPages          （空白页插入）
3. 提取图片宽度百分比         （辅助工具，可选）
4. 显示表格列宽百分比         （辅助工具，可选）
```

**为什么续表标签必须先于空白页插入？**

`AddContinuedTableLabels` 会分割跨页表格并在分割处插入"（续表x.y）"段落，这会改变页面布局（行数、分页位置均会变化）。如果先运行 `InsertBlankPages` 插入空白页，再分割表格，之前插入的空白页位置可能因布局变化而错位。

正确做法是先分割表格、稳定页面布局，再根据最终布局插入空白页。

### 9.3 各宏详解

#### ContinuedTable.vba — 续表标签

**功能：** 检测跨页表格，在分页处分割表格，插入"（续表x.y）"标签。

**处理逻辑：**

1. 更新所有域
2. 遍历所有表格，检测是否跨页
3. 找到跨页表格的分页行
4. 在分页行处分割表格
5. 在两个表格之间插入"（续表x.y）"段落（右对齐，宋体五号）
6. 验证续表文字是否在新页面，否则强制分页
7. 每次只分割一处，然后从头重新扫描（确保分页位置准确）
8. 重复直到无跨页表格

**跳过条件：**

- 图片布局表格（含 InlineShape 图片的表格）
- 无法提取表编号的表格

**使用：** F5 运行 `AddContinuedTableLabels`

#### InsertBlankPages.vba — 空白页插入

**功能：** 确保每个一级标题（Heading1）从奇数页开始。

**处理逻辑：**

1. 更新所有域
2. 遍历所有段落，找到 Heading1 标题
3. 检查标题所在页码是否为偶数
4. 若为偶数，在前一节末尾（分节符之前）插入分页符
5. 更新域后重新检测（插入空白页可能影响后续页码）
6. 重复直到无 Heading1 在偶数页

**使用：** F5 运行 `InsertBlankPages`

> 此宏适用于双面打印的论文，确保每章从右手页（奇数页）开始。

#### GetImageWidth.vba — 图片宽度提取

**功能：** 提取文档中所有图片相对于页面内容宽度的百分比，用于回填 QMD 中的 `width` 属性。

**输出格式：**

```
序号    宽度      fig标签      描述(AltText)
--------------------------------------------------------------------------------
1       width="80%"   fig-framework    研究框架
2       width="100%"  fig-lstm         LSTM架构
```

底部还提供可直接粘贴到 QMD 的 `width` 属性片段。

**使用场景：** 在 Word 中手动调整图片大小后，运行此宏提取宽度百分比，回填到 QMD 源文件中，使下次编译时图片自动使用正确宽度。

**使用：** F5 运行 `提取图片宽度百分比`

#### GetTableColWidth.vba — 表格列宽查看

**功能：** 显示光标所在表格的各列宽度百分比。

**使用：**

1. 将光标放在表格内
2. F5 运行 `显示表格列宽百分比`
3. 结果复制到剪贴板

#### FieldCodeToText.vba — 域代码转文本

**功能：** 将所有域代码转为纯文本显示。用于调试和检查域代码内容。

**使用：** F5 运行 `fieldcodetotext`

> ⚠️ 此宏会**不可逆地**将域转为纯文本，仅用于调试，不要在正式文档上运行。建议先备份。

## 10. 精确排版与手动调整

本项目通过自动化处理覆盖了绝大部分格式要求，但由于 Quarto/Pandoc 生成 DOCX 的固有局限，以及 Word 排版引擎的特殊性，**部分精确排版仍需在编译后的 DOCX 中手动完成**。

### 需要手动调整的项目

| 项目         | 说明                       | 调整方法                                               |
| ------------ | -------------------------- | ------------------------------------------------------ |
| 目录更新     | 编译后目录内容可能不完整   | Word 中全选（Ctrl+A）→ 右键 → 更新域 → 更新整个目录 |
| 交叉引用更新 | 部分引用编号可能未刷新     | Ctrl+A → F9 刷新所有域                                |
| 图片精确位置 | 图片与题注的间距可能不理想 | 手动微调图片段落间距                                   |
| 表格行高     | 某些行高可能需要微调       | 选中行 → 表格属性 → 行 → 指定高度                   |
| 页眉内容     | 部分节的页眉可能需要微调   | 双击页眉区域编辑                                       |
| 续表标签     | 必须通过 VBA 宏处理        | 运行 `AddContinuedTableLabels`                       |
| 奇数页起始   | 必须通过 VBA 宏处理        | 运行 `InsertBlankPages`                              |

### 推荐工作流

```
1. 编辑 thesis.qmd
2. 运行编译脚本
3. 在 Word 中打开 output/thesis.docx
4. Ctrl+A → F9 更新所有域
5. 运行 VBA 宏（续表 → 空白页）
6. 手动微调不满意的格式
7. 最终检查
```

> ⚠️ 每次 recompile 后，手动调整会丢失。建议在 QMD 中尽可能完善内容和格式，减少手动调整量。对于必须手动调整的项目，可考虑将其自动化为后处理步骤。

## 11. 后处理流水线详解

### 11.1 流程总览

`thesis_post_process.py` 执行以下步骤（顺序有依赖，不可随意调整）：

```
1.  页边距设置
2.  目录样式处理
2.5 摘要/Abstract 移至目录之前
2.6 排除摘要/Abstract 出现在 TOC 中
2.7 Abstract 标题加粗
2.8 关键词黑体字体
3.  非正文章节禁用编号
4.  公式处理（编号、右对齐、eqArr结构）
4.  图片样式和题注
5.  表格样式和题注（三线表）
5.1 创新成果自评表全边框
5.2 QCA 符号处理
6.  交叉引用更新
6.1 移除交叉引用前后多余空格
7.  页眉页脚（奇偶页不同、章节标题页眉）
8.  定理环境交叉引用
8.1 移除定理引用前后多余空格
8.5 算法样式
8.6 全局去除行内公式前后多余空格
9.  含公式段落行间距
9.1 公式后续接段落处理
10. 参考文献样式
11. 科研成果章节样式
11.5 横向页面表格
11.6 脚注处理（①②③，每页重启）
12. 封面页合并
    → 保存 XML
    → 封面页文件操作（header/footer、rels）
    → 重新打包 DOCX
```

### 11.2 步骤依赖说明

- **步骤 2 必须在步骤 7 之前**：目录样式处理需在页眉页脚之前，因为页眉页脚会创建新的 sectPr
- **步骤 4（公式）必须在步骤 6（交叉引用）之前**：交叉引用依赖公式编号映射
- **步骤 7（页眉页脚）必须在步骤 11.5（横向表格）之前**：横向表格需要完整的 sectPr 属性
- **步骤 11.6（脚注）在步骤 11.5 之后**：脚注注入到所有 sectPr，需在 sectPr 完整后执行
- **步骤 12（封面合并）在最后**：封面合并涉及文件系统操作，需在所有 XML 修改完成后执行

### 11.3 模块职责

| 模块                         | 职责                                                      |
| ---------------------------- | --------------------------------------------------------- |
| `config.py`                | 集中配置：页边距、字体、样式ID、排除章节列表、公式制表位  |
| `utils.py`                 | DOCX 解压/打包、命名空间注册、mc:Ignorable 修复、XML 读写 |
| `page_style.py`            | 页边距设置                                                |
| `equation_style.py`        | 公式编号、eqArr 结构、右对齐制表位、# 号隐藏              |
| `figure_style.py`          | 图片样式、英文标题、交叉引用                              |
| `table_style.py`           | 三线表样式、英文标题、表注、自定义边框、列对齐            |
| `header_footer.py`         | 页眉页脚（奇偶页不同、章节标题、论文题目页眉）            |
| `footnote_style.py`        | 脚注①②③编号、每页重启、settings.xml/sectPr 配置        |
| `symbol_style.py`          | QCA 符号处理（⬤⊗●⊘ 大小区分、字体统一）               |
| `title_pages.py`           | 封面页合并（内容插入 + 文件复制 + rels 更新）             |
| `toc_style.py`             | 目录样式（TOC 域代码修改、书签范围）                      |
| `abstract_style.py`        | 摘要移至目录前、Abstract 加粗、关键词黑体、自评表边框     |
| `heading_style.py`         | 非正文章节禁用编号                                        |
| `cross_reference.py`       | 交叉引用更新、空格清理                                    |
| `theorem_style.py`         | 定理环境交叉引用                                          |
| `algorithm_style.py`       | 算法样式                                                  |
| `landscape_table.py`       | 横向页面表格                                              |
| `bibliography.py`          | 参考文献样式                                              |
| `paragraph_style.py`       | 含公式段落行间距、行内公式空格清理                        |
| `equation_continuation.py` | 公式后续接段落                                            |

### 11.4 OOXML 注意事项

后处理直接操作 DOCX 内部 XML，需注意以下 OOXML 约束：

**mc:Ignorable 命名空间修复**

Python `ElementTree` 在写入 XML 时只声明实际使用的命名空间前缀，但 `mc:Ignorable` 属性以字符串引用前缀（如 `"w14 w15 w16se"`）。如果这些前缀没有 `xmlns` 声明，Word 会报"无法读取的内容"。

`utils.py` 中的 `fix_mc_ignorable_namespaces()` 函数在打包前扫描所有 XML 文件，确保 `mc:Ignorable` 引用的每个前缀都有对应的 `xmlns` 声明。

**sectPr 子元素顺序**

OOXML schema 对 `CT_SectPr` 的子元素顺序有严格要求：

```
headerReference* → footerReference* → endnotePr? → footnotePr? → type? → pgSz? → ...
```

后处理在插入 `footnotePr`、`type` 等元素时，会查找正确的锚点位置，确保顺序合规。

**settings.xml 中 footnotePr 的位置**

`w:footnotePr` 在 `w:settings` 下必须出现在 `w:endnotePr`、`w:compat`、`w:rsids` 等元素之前。

## 12. Lua 过滤器说明

### eq_number.lua — 公式分章编号

- 监听 `Header` 事件，一级标题时章号 +1，公式计数器归零
- 监听 `Para` 事件，检测行间公式（`DisplayMath`），提取 `{#eq-xxx}` 标签
- 为公式生成编号（如 `(2-1)`），包裹在 `equation-style` 自定义样式的 Div 中
- 监听 `Cite` 事件，将 `@eq-xxx` 引用转为超链接

### next_page.lua — 一级标题分节符

- 在除第一个一级标题外的所有一级标题前，插入 `w:sectPr`（分节符，nextPage 类型）
- 确保每章从新页开始

### landscape.lua — 横向表格

- 检测带有 `.landscape` class 的 Div，在 Div 前后插入分节符实现页面旋转
- 交换 `pgSz` 的 `w` 和 `h` 属性实现页面旋转

> **注意：** 本项目 `thesis.qmd` 中横向表格实际使用 `<!-- landscape: 关键词 -->` HTML 注释标记，由 Python 后处理脚本 `landscape_table.py` 从 QMD 源文件读取并处理，不依赖此 Lua 过滤器。如需使用 Lua 方式，需在 QMD 中用 `::: {.landscape}` Div 包裹表格。

### theorem_env.lua — 定理环境

- 支持环境：definition / lemma / theorem / corollary / proposition / example / remark / assumption / proof
- 分章编号（如 定义2.1、定理3.1）
- 支持 `name="自定义名称"` 属性
- 证明环境不编号，自动添加 QED 符号（∎）
- 交叉引用：`@def-xxx` → 定义2.1

## 13. 模板文件

### heu_thesis_style.docx

Quarto 参考模板，定义了论文中使用的段落样式和字符样式。后处理脚本通过样式 ID（如 `af3`、`afa`、`afc`）引用这些样式。

**关键样式映射（定义在 `config.py` 中）：**

| 用途     | 样式ID             | 样式名称     |
| -------- | ------------------ | ------------ |
| 一级标题 | `1`              | Heading 1    |
| 图片题注 | `af3`            | ImageCaption |
| 图片段落 | `af6`            | 图           |
| 表格题注 | `af3`            | 题注         |
| 表格内容 | `af5`            | 表           |
| 三线表   | `aff0`           | 三线表       |
| 参考文献 | `af7`            | 参考文献     |
| 页眉     | `ae`             | 页眉         |
| 页脚     | `af0`            | 页脚         |
| 公式     | `equation-style` | 公式         |
| 脚注文字 | `afa`            | 脚注文字     |
| 脚注引用 | `afc`            | 脚注引用     |

> 如需适配其他学校模板，需修改此文件和 `config.py` 中的样式映射。

### thesis_title_pages.docx

封面页源文件。包含封面、独创声明、学位论文版权使用授权书等页面。编译时自动合并到正文之前。

## 14. 常见问题与排错

### Word 报"无法读取的内容"

**原因：** `mc:Ignorable` 属性引用了未声明的命名空间前缀，或 `sectPr` 子元素顺序不符合 OOXML schema。

**排查：**

1. 确认 `utils.py` 中的 `fix_mc_ignorable_namespaces()` 已在 `pack_docx()` 中被调用
2. 确认 `_MC_PREFIX_TO_URI` 映射包含所有 `mc:Ignorable` 引用的前缀
3. 检查 `sectPr` 中 `footnotePr`、`type` 等元素的插入位置是否正确

### 脚注不显示带圈数字

**原因：** `footnotePr` 配置未正确写入 `settings.xml` 或 `sectPr`。

**排查：**

1. 确认 `settings.xml` 中 `w:footnotePr` 包含 `w:numFmt w:val="decimalEnclosedCircle"`
2. 确认 `sectPr` 中也注入了 `footnotePr`（节级设置覆盖全局）
3. 确认 `footnotePr` 不包含 `<w:footnote>` 子元素（这会导致"无法读取的内容"）

### 交叉引用编号未更新

**解决：** Word 中 Ctrl+A → F9 刷新所有域。或右键目录 → 更新域 → 更新整个目录。

### 公式编号格式异常

**排查：**

1. 确认 `filters/eq_number.lua` 在 YAML 的 `filters` 列表中
2. 确认公式使用 `{#eq-xxx}` 标签
3. 确认引用使用 `@eq-xxx` 语法

### 编译脚本报错

**常见原因：**

- `quarto: command not found` → Quarto 未安装或不在 PATH 中
- `python3: command not found` → Python 未安装（Windows 用 `python`）
- `ModuleNotFoundError` → 确认在项目根目录运行脚本

### QCA 符号大小无区分

**排查：** 确认 YAML 头中 `qca-symbols: true` 已启用（非注释行）。

## 15. 扩展指南

### 添加新的后处理步骤

1. 在 `scripts/docx_processor/` 下创建新模块（如 `new_feature.py`）
2. 实现处理函数，接受 `root`（XML 根元素）和其他必要参数
3. 在 `thesis_post_process.py` 中导入并调用，注意步骤顺序
4. 如需从 QMD 读取配置，参考 `config.py` 中的 `read_qmd_yaml_bool()`
5. 如需从 QMD 读取内容，参考 `table_style.py` 中的 `load_*_from_qmd()` 系列函数

### 适配其他学校模板

1. 替换 `docx_template/heu_thesis_style.docx` 为目标学校的模板
2. 修改 `config.py` 中的样式 ID 映射（用 Word 打开模板查看样式 ID）
3. 修改 `config.py` 中的页边距、排除章节列表等配置
4. 修改 `header_footer.py` 中的页眉页脚规则
5. 替换 `thesis_title_pages.docx` 为目标学校的封面页
6. 修改 `abstract_style.py` 中的摘要/关键词处理逻辑

### 添加新的 YAML 配置项

1. 在 `config.py` 中使用 `read_qmd_yaml_bool()` 读取布尔配置
2. 在 `thesis_post_process.py` 中根据配置条件调用处理函数
3. 在 QMD 的 YAML 头中添加配置项

### 添加新的 Lua 过滤器

1. 在 `filters/` 下创建 `.lua` 文件
2. 在 YAML 的 `filters` 列表中添加（注意顺序）
3. Lua 过滤器运行在 Quarto 渲染阶段，适合内容转换（编号、标签、HTML 注入）
4. 需要精细 XML 控制的功能应放在 Python 后处理阶段
