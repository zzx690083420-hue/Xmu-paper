# 厦门大学论文格式工具 — CLAUDE.md

## 项目概述

本工具是厦门大学学位论文格式自动化处理程序，基于《厦门大学管理学院MBA学位论文格式规范》（2025年11月版）开发。上传 `.docx` 文件后，自动修正页面设置、字体、标题样式、页眉页脚、目录等格式，输出符合规范的 Word 文档。

## 技术栈

- **后端**：Python + Flask (`app.py`)
- **格式处理核心**：`formatter.py`（基于 `python-docx` + 直接操作 OOXML）
- **前端**：单页 HTML（`templates/index.html`）
- **依赖**：`flask>=3.0.0`、`python-docx>=1.1.0`

## 启动方式

```bash
bash start.sh          # 自动创建虚拟环境、安装依赖、启动服务
# 或手动：
source venv/bin/activate
python3 app.py         # 访问 http://localhost:5001
```

## 文件结构

```
formatter.py        # 核心格式处理逻辑，所有格式规则在此
app.py              # Flask Web 服务，提供 /format 和 /preview 接口
templates/index.html  # 前端页面
requirements.txt    # 依赖列表
start.sh            # 一键启动脚本
```

## formatter.py 架构

### 入口函数
`format_thesis(input_path, output_path, options)` — 按顺序调用各子函数，返回修改列表。

### 主要子函数（按调用顺序）
| 函数 | 作用 |
|------|------|
| `update_page_setup(doc)` | 页边距、纸张大小、页眉页脚距离 |
| `update_styles(doc)` | 修正所有样式定义（Normal/Heading/toc/Caption 等） |
| `fix_heading_direct_format(doc)` | 清除标题段落直接格式，强制居中/左对齐，特殊节标题居中 |
| `auto_detect_headings(doc)` | 按文字模式自动识别并应用 Heading 样式（默认关闭） |
| `fix_body_text_fonts(doc)` | 修正正文字体、首行缩进、行距 |
| `fix_captions_and_tables(doc)` | 题注格式（加粗）、跨页表格防护 |
| `regenerate_toc(doc)` | 重新生成中英文双语目录（默认关闭） |
| `update_headers(doc, title)` | 为每章建立独立节，设置奇偶页眉 |
| `update_footer_page_numbers(doc)` | 设置页码（封面/目录罗马数字，正文阿拉伯数字从1开始） |

### 关键辅助函数
- `_get_chapter_heading_style(doc)` — 检测文档实际使用的章节样式名（Heading 1/2 等）
- `_find_post_toc_idx(doc)` — 返回目录结束后第一个段落的索引，用于限定章节处理范围
- `_insert_chapter_section_breaks(doc, chapter_style)` — 在各章前插入分节符，返回 `{section_idx: chapter_title}` 映射
- `_make_toc_field_paragraphs(doc, ...)` — 生成静态英文目录段落（含 PAGEREF 书签页码）
- `_add_bookmark(p_elem, id, name)` — 在段落 `w:pPr` 之后插入书签（OOXML 规范要求）
- `_find_post_toc_idx(doc)` / `_find_first_chapter_idx_after_toc` — TOC 边界检测

## 格式规范常量（`formatter.py` 顶部）

| 常量 | 值 | 说明 |
|------|-----|------|
| `PAGE_MARGIN_LEFT/RIGHT` | 2.8 cm | 左右页边距 |
| `PAGE_MARGIN_TOP/BOTTOM` | 2.54 cm | 上下页边距 |
| `FONT_CHINESE_BODY` | 宋体 | 正文中文字体 |
| `FONT_CHINESE_HEADING` | 黑体 | 标题中文字体 |
| `FONT_ENGLISH` | Times New Roman | 英文字体（全文统一） |
| `SIZE_NORMAL` | 12pt（小四） | 正文字号 |
| `FIRST_LINE_INDENT` | 24pt | 正文首行缩进（2字符） |

## 重要 OOXML 规则（避免踩坑）

- **`w:bookmarkStart` 必须在 `w:pPr` 之后**：放在 `w:pPr` 前会导致 Word 计算 PAGEREF 页码错误
- **内联 `w:sectPr` 定义的是它所在段落的节**（该节在此段落结束），不是下一节
- **`w:pgNumType w:start="1"`** 写入某节的 sectPr，使该节页码从 1 开始
- **目录样式 `numPr` 会引入额外缩进**：需同时从样式定义和段落直接格式中移除
- **中文 TOC 段落样式名不固定**：可能是 "toc 1"、"TOC 1"、"目录 1" 等，处理时用正则扫描所有样式

## 章节处理逻辑

1. `_find_post_toc_idx(doc)` 找到目录结束位置
2. 只处理目录之后、匹配 `HEADING1_PATTERNS` 的标题段落
3. `_insert_chapter_section_breaks` 在每章前插入 `nextPage` 分节符（含第一章）
4. 每章独立节 → 独立页眉（StyleRef 域自动显示章节标题）
5. 第一章所在节设置 `pgNumType w:start=1`，使正文页码从 1 开始

## Web API

| 路由 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 前端页面 |
| `/format` | POST | 上传 `.docx`，返回格式化后的文件下载 |
| `/preview` | POST | 上传 `.docx`，返回修改列表 JSON（不返回文件） |

`/format` 和 `/preview` 均接受以下 form 参数（均为 `'true'/'false'` 字符串）：
`fix_page_setup`、`fix_styles`、`fix_heading_direct`、`auto_detect_headings`、`fix_body_fonts`、`add_headers`、`add_page_numbers`、`regenerate_toc`、`thesis_title`

## 开发注意事项

- 格式规则改动只需修改 `formatter.py`，不需要动 `app.py` 或前端
- 测试时直接用 `python3 formatter.py` 不可行（无 CLI 入口），需通过 Web 界面上传文档
- `python-docx` 的高层 API（`para.paragraph_format.left_indent` 等）有时不能覆盖直接格式，优先直接操作 `para._p` XML 元素
- 所有章节处理必须在目录之后开始（用 `_find_post_toc_idx` 限定），避免封面/摘要的 Heading 样式段落被误处理
