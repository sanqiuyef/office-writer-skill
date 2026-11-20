---
name: office-writer
description: 用 45 个独立 CLI 入口创建、编辑、读取、转换和验收 Word/Excel/PPT/PDF：写报告/公文/论文/简历/合同、Markdown 转 Word、Word 批注/域/属性/修订/脚注/内容控件、Excel 图表/条件格式/数据验证/超级表/定义名称/透视表、PPT 超链接/几何/裁剪/动画/切换/批注/SmartArt/形状、公式重算、PPT 配色/布局/清理/框架图、套模板填充 {{占位符}}、导出 PDF/打印、批量编辑与活动 Office 文档。转换/渲染/重算支持 dsoffice/Office/WPS COM 回退。触发：写 Word、生成 docx、写报告、写公文、写论文、写简历、写合同、改稿、追加段落、写 Excel、生成表格、公式重算、写 PPT、做演示文稿、PPT 美化、PPT 配色、PPT 优化布局、清理占位符、生成框架图、打开 PPT 画形状、写 PDF、Markdown 转 Word、套模板、填充占位符、导出 PDF、转 PDF、渲染页面、检查排版、读取文档、批量编辑文档、活动 Office 文档。
---

# Office Writer

一个能力一个入口，共 45 个脚本：其中 32 个与 mcp-office-writer 0.6.0 的 32 个工具 1:1 对应，另加 13 个本地入口：`replace_open_document_content.py`（活动文档内容替换）、`word_manage_comments.py`（批注增/列）、`word_set_properties.py`（核心属性）、`excel_add_chart.py`（已有文件加图表）、`ppt_set_hyperlinks.py`（已有文件设超链接），以及 8 个 office-native 原生入口（Word 修订/脚注/内容控件、PPT 动画切换批注/SmartArt/形状绘制、Excel 透视表/打印）。默认由 `scripts/office_engine/`（Python）执行；C# `tools/office-native/`（Open XML SDK + Office COM）作为可选后端，并自动承接 7 个 Python 未实现的入口（Word 修订/脚注/内容控件、PPT 动画切换批注/SmartArt、Excel 透视表/打印）。本 skill 只做 CLI 适配、JSON 契约和验收辅助，不启动 MCP 服务。

## 触发条件

用户要求创建、修改、读取、转换或验收办公文档时使用：

- Word：写报告/公文/论文/简历/合同、套模板、Markdown 转 Word、追加内容、目录、页眉页脚、封面。
- Excel：多工作表、表头/格式/合并/冻结/筛选、公式写入与重算。
- PowerPoint：新建演示文稿、追加幻灯片、配色、布局优化、清理占位符、分析模板、生成框架图/流程图、在已打开的演示里直接画形状。
- PDF：新建 PDF、Office 转 PDF、逐页渲染 PNG、结构化读取、字号/字体验收。
- 结构化编辑：按 `/body/p[1]`、`/Sheet1/A1`、`/slide[1]/shape[2]` 路径原子批量编辑。
- 活动文档：Word/WPS、Excel/WPS 表格、PowerPoint 已打开文档的实时追加。

## 修改已有文件：禁止复写（默认工作方式，脚本强制）

已有 Office 二进制文件（.doc/.docx/.ppt/.pptx/.xls/.xlsx/.docm 等）的修改，**禁止用 `create_*` 重生成整份文件覆盖**，只能在原文件基础上增删改查：

1. **文件已打开** → 活动文档入口原地改，由应用自己保存：
   - Word：`replace_open_document_content`（按标题替换整节 / 全文查找替换）、`edit_open_word_document`（追加内容）
   - PowerPoint：`replace_open_document_content`（形状文本查找替换）
   - Excel：`replace_open_document_content`（单元格查找替换）

   只有应用自己写回（COM），才能保留 MathType/OLE 公式、修订、批注、域、嵌入对象与手工样式。
2. **文件未打开** → 文件级编辑入口（`batch_edit_office_document`、`edit_word_document`、`edit_powerpoint`、`edit_excel_workbook`）修改原文件。这些入口在写回前自动生成 `*.bak-YYYYmmdd-HHMMSS` 备份，备份路径在 `data.backup_path` 返回。
3. **只有两种情况可以新建或复写**：
   - 目标是不存在的新文件，或空白文件（0 字节）；
   - 用户明确授意重生成或整体覆盖（例如"重新生成一份""按模板重排"），且必须显式传 `--if-exists overwrite --confirm-overwrite`。

覆盖闸门由脚本执行，不依赖模型记忆：所有写文件的入口默认 `--if-exists error`，遇到已存在且非空的目标直接退出码 2，错误信息说明三条出路（改用编辑入口 / 加 `rename` 保留旧文件 / 确认后覆盖）。`create_word_document` / `create_word_from_markdown` / `create_powerpoint` / `create_excel_workbook` 等重生成入口只用于第 3 条；其余一律走 1 或 2。拿不准时先问用户，**不复写**。

```bash
# 目标已存在且非空：默认拒绝，退出码 2
python -X utf8 scripts/create_word_document.py --output-path 报告.docx --params '{...}'

# 保留旧文件：自动改名 报告 (1).docx
... --if-exists rename

# 用户明确授意整体覆盖：两个开关缺一不可，覆盖前自动写 .bak 备份
... --if-exists overwrite --confirm-overwrite
```

## 何时不要用

- 只想“想文案/改措辞”，不需要生成或修改文档 → 直接写文本或使用写作类 skill。
- 只做 Excel 数据排序、透视、统计建模 → 用 pandas/openpyxl 脚本；本 skill 负责版式与写入。
- 只做 PDF 拆分/合并/加密/OCR 校对 → 用专门的 PDF 工具；本 skill 的 OCR 仅用于读取扫描件。
- 需要生成示意图/海报/插画 → 用绘图或生图工具；`ppt_create_framework_diagram` 只画阶段+节点+箭头。
- 需要启动 Word/WPS/Excel/PowerPoint → 活动文档入口只连接现有窗口，不启动应用；转换/渲染/重算的 COM 回退是例外：它用 `DispatchEx` 启动隐藏独立实例，完成后 Quit。
- 非 Windows 环境需要 COM 活动文档或 COM 转换 → 用文件级入口并安装 dsoffice；COM 回退会明确报不可用。
- 未构建 office-native：默认 Python 入口不受影响；7 个 C# 专属入口（修订/脚注/内容控件/动画/切换/批注/SmartArt/透视表/打印）返回退出码 3，需先构建，或人工在 Office 中处理。

## 能力面与边界

Word / Excel / PPT 的逐项功能面与覆盖状态见三张矩阵；标注「未覆盖」的能力不要用现有入口伪造。

| 应用 | 能力矩阵 | 已覆盖重点 | 主要缺口（未封装） |
|:---|:---|:---|:---|
| Word | `references/capabilities_word.md` | 结构/样式/表格/图片/页眉页脚、域、TOC、超链接、批注、文档属性、修订、脚注、内容控件、书签、分栏/页面边框/水印、模板填充 | 邮件合并、比较合并、打印、浮动图片/裁剪、图表/SmartArt/OLE/公式创建 |
| Excel | `references/capabilities_excel.md` | 多表/样式/行高列宽、8 类图表、条件格式、数据验证、超级表、定义名称、批注、超链接、图片、页面设置、真实排序、工作表保护、公式重算、透视表、PDF/打印 | 形状、迷你图/切片器、宏、加密、Power Query、工作簿结构保护 |
| PPT | `references/capabilities_ppt.md` | 版式/占位符/文本/真列表/表格/图表/图片裁剪/形状/连接线/框架图/超链接/备注/页脚页码/节/放映设置/配色/布局审计、动画、切换、批注、SmartArt | 母版、自定义放映、组合形状、媒体/OLE、单页导出/视频 |

## office-native 后端（默认）

**全部 45 个入口都有 C# 实现**（Open XML SDK 文件级 + Office COM 独占能力），源码、命令与边界见 `tools/office-native/README.md`。默认执行策略是 **Python 优先**：`scripts/office_engine/` 有实现的入口走 Python；7 个 Python 未实现的入口（Word 修订/脚注/内容控件、PPT 动画切换批注/SmartArt、Excel 透视表/打印）自动走 C#。

- 构建：`pwsh -File tools/office-native/build.ps1`（需 .NET SDK 8.0+）；产物 `tools/office-native/dist/office-native.exe`；也可用 `OFFICE_NATIVE_BIN` 指定。
- 切换：默认 auto（Python 优先，缺口自动走 C#）；`OFFICE_NATIVE_ENGINE=csharp` 强制 C#；`OFFICE_NATIVE_ENGINE=python` 强制 Python（C# 专属入口没有 Python 实现，仍会调用 C#）。
- 未构建时只有 C# 专属入口返回退出码 3 并给出构建命令；Python 入口照常工作。
- 要求 Windows + 对应 Office（SmartArt 需要带窗口打开，命令会自动处理）；修改前自动生成 `.bak-*`；COM 转换用隔离实例，不附着用户窗口。

## 能力→脚本索引

所有脚本在 `scripts/`，用同一解释器（Python ≥ 3.10）运行；通用前缀：

```bash
python -X utf8 scripts/<脚本名>.py --help
```

所有入口都支持 `--params '<JSON 对象>'` 与 `--params-file <路径>`；复杂模型参数放在 JSON 中，由 Pydantic `model_validate` 校验。显式 flag 优先于同名 `--params` 键。

| 脚本 | 能力 | 关键参数 |
|:---|:---|:---|
| `scripts/create_word_document.py` | 创建 Word；模板/场景/封面/页眉页脚 | `--output-path`、`--template`、`--scene`、`--params`（`content`、`page_setup`、`header_footer`、`default_font`、`cover_config`） |
| `scripts/edit_word_document.py` | 向已有 Word 追加内容块 | `--file-path`、`--params`（`append_sections`、`default_font`） |
| `scripts/create_word_from_markdown.py` | Markdown 长文转 Word | `--source-path`、`--output-path`、`--template`、`--scene` |
| `scripts/create_excel_workbook.py` | 创建 Excel 多工作表 | `--output-path`、`--template`、`--params`（`sheets`） |
| `scripts/edit_excel_workbook.py` | 向已有 Excel 追加工作表 | `--file-path`、`--params`（`append_sheets`） |
| `scripts/create_powerpoint.py` | 创建 PPT | `--output-path`、`--template`、`--params`（`slides`） |
| `scripts/edit_powerpoint.py` | 向已有 PPT 追加幻灯片 | `--file-path`、`--params`（`append_slides`） |
| `scripts/create_pdf_document.py` | 创建 PDF | `--output-path`、`--params`（`content`、`page_setup`、`font_config`） |
| `scripts/convert_to_pdf.py` | Office → PDF（dsoffice → Office/WPS COM） | `--input-path`、`--output-path`、`--sheet`、`--if-exists`；返回 `backend`、`com_engine`、`fallback_from[]`、`font_substitutions[]` |
| `scripts/render_document_pages.py` | 逐页渲染 PNG（dsoffice；缺失时 COM→PyMuPDF） | `--input-path`、`--output-dir`、`--pages`、`--dpi`、`--sheet`、`--cell-range`；`page_count` 按实际产物修正 |
| `scripts/recalculate_excel_formulas.py` | 重算公式并写缓存值（dsoffice；缺失时 Excel/WPS COM） | `--input-path`、`--output-path`、`--if-exists` |
| `scripts/get_open_office_documents.py` | 列出活动 Office 文档 | 无 |
| `scripts/edit_open_word_document.py` | 编辑活动 Word/WPS 文档 | `--file-path`、`--params`（`append_sections`） |
| `scripts/edit_open_excel_workbook.py` | 编辑活动 Excel 工作簿 | `--file-path`、`--params`（`append_sheets`） |
| `scripts/edit_open_powerpoint.py` | 编辑活动 PowerPoint | `--file-path`、`--params`（`append_slides`） |
| `scripts/replace_open_document_content.py` | **活动文档内容替换（不重生成）**：Word 按标题替换整节/全文查找替换；PPT、Excel 查找替换 | `--file-path`、`--params`（`replace_sections`、`replace_texts`） |
|  | ⚠️ PowerPoint 活动文档保存失败时，C# 后端会删除本次新增幻灯片并返回退出码 3；`OFFICE_NATIVE_ENGINE=python` 的旧引擎只提示用户在 PowerPoint 中按 Ctrl+S。 |  |
| `scripts/ppt_list_color_schemes.py` | 列出 12 套 PPT 配色 | 无 |
| `scripts/ppt_apply_color_scheme.py` | 应用 PPT 配色 | `--file-path`、`--scheme-name` |
| `scripts/ppt_optimize_slide_layout.py` | 审计 PPT 越界/重叠，返回 `warnings[]`；`--reflow` 自动重排并写回 | `--file-path`、`--reflow`、`--fix-spacing`（旧的纵向间距优化） |
| `scripts/ppt_clean_unused_placeholders.py` | 清理空白占位符 | `--file-path` |
| `scripts/ppt_analyze_template.py` | 分析 PPT 模板 | `--file-path` |
| `scripts/ppt_create_framework_diagram.py` | 生成 PPT 框架图 | `--output-path`、`--params`（`diagram`） |
| `scripts/ppt_add_flowchart_slide.py` | 活动 PPT 追加流程图 | `--file-path`、`--params`（`diagram`） |
| `scripts/read_word_document.py` | 读取 Word | `--file-path`、`--max-chars` |
| `scripts/read_excel_workbook.py` | 读取 Excel | `--file-path`、`--max-rows-per-sheet`、`--max-chars` |
| `scripts/read_powerpoint.py` | 读取 PPT | `--file-path`、`--max-chars` |
| `scripts/read_pdf_document.py` | 读取 PDF 文本 | `--file-path`、`--max-chars` |
| `scripts/read_pdf_structured.py` | 读取 PDF 结构化内容 | `--file-path`、`--include-blocks`、`--extract-images` |
| `scripts/read_txt_document.py` | 读取纯文本 | `--file-path`、`--max-chars` |
| `scripts/inspect_office_document.py` | 读取节点树/稳定路径 | `--file-path`、`--max-nodes` |
| `scripts/validate_word_document.py` | Word 质量门控：结构 + 字体安装/标题编号/PAGE/TOC/表格列/空段落/中英混排 | `--file-path`；`checks[].id` 稳定，`errors>0` 退出 1 |
| `scripts/merge_office_template.py` | 填充 `{{占位符}}` | `--template-path`、`--output-path`、`--params`（`data`） |
| `scripts/batch_edit_office_document.py` | 原子批量编辑 | `--file-path`、`--output-path`、`--no-backup`、`--params`（`operations`） |
| `scripts/word_manage_comments.py` | Word 批注：按文字/段落锚定新增、列出 | `--file-path`、`--list`、`--params`（`add`：`text`/`author`/`anchor_text`/`paragraph_index`） |
| `scripts/word_set_properties.py` | Word 核心属性：读取/写入 | `--file-path`、`--list`、`--params`（`properties`） |
| `scripts/excel_add_chart.py` | 已有 Excel 插入图表（8 类，锚点/宽高） | `--file-path`、`--sheet`、`--params`（`chart`） |
| `scripts/ppt_set_hyperlinks.py` | 已有 PPT 设形状/文字 run 超链接 | `--file-path`、`--params`（`links`：`shape_index`/`shape_text`/`text_match`） |
| `scripts/word_manage_revisions.py` | Word 修订：列/接受/拒绝/标记插入与删除 | `--file-path`、`--action`、`--params`（`text`/`anchor_text`/`paragraph_index`）；C# 专属（Python 无实现） |
| `scripts/word_manage_footnotes.py` | Word 脚注：新增/列出 | `--file-path`、`--action`、`--params`（`text`/`anchor_text`）；C# 专属（Python 无实现） |
| `scripts/word_manage_content_controls.py` | Word 内容控件：文本/日期/下拉/复选框增删改查 | `--file-path`、`--action`、`--tag`、`--params`；C# 专属（Python 无实现） |
| `scripts/ppt_manage_effects.py` | PPT 动画/切换/批注 | `--file-path`、`--action`、`--slide`、`--params`；C# 专属（Python 无实现） |
| `scripts/ppt_manage_smartart.py` | PPT SmartArt：列布局/插入 | `--file-path`、`--action`、`--slide`、`--params`（`layout`/`texts`）；C# 专属（Python 无实现） |
| `scripts/ppt_manage_shapes.py` | PPT 形状绘制：活动文档/文件级加形状、列形状 | `--file-path`、`--action`（add/list）、`--slide`、`--params`（`shape`）；默认 Python，文件已打开走活动文档 COM，未打开走文件级并自动备份；C# 可选 |
| `scripts/excel_manage_pivots.py` | Excel 透视表：创建/列出/刷新 | `--file-path`、`--action`、`--sheet`、`--params`；C# 专属（Python 无实现） |
| `scripts/excel_print.py` | Excel 导出 PDF / 打印 | `--file-path`、`--action`、`--sheet`、`--output`、`--params`；C# 专属（Python 无实现） |

`SectionItem.type` 另支持 `hyperlink`（`HyperlinkDef`）与 `field`（`FieldDef`：PAGE/NUMPAGES/SECTIONPAGES/DATE/TIME/FILENAME/AUTHOR/TITLE/CUSTOM）。`SheetDef` 另支持 `charts`/`conditional_formats`/`data_validations`/`tables`/`defined_names`/`hyperlinks`/`comments`/`images`/`page_setup`/`row_heights`/`sheet_state`；`SlideTableDef`/`ChartDef` 支持 x/y/width/height，`SlideImageDef` 支持四边裁剪与超链接。参数细节与 JSON 模型见 `references/cli_args.md`，全部字段（含默认值/类型/说明）见 `references/models_word.md`、`models_excel.md`、`models_ppt.md`、`models_pdf_common.md`；模型定义源在 `scripts/office_engine/models.py`。转换/渲染/重算的后端顺序用 `OFFICE_CONVERT_BACKEND=auto|dsoffice|com|none` 控制：默认 auto 依次尝试 dsoffice、Microsoft Office COM、WPS COM；dsoffice 安装与 COM 排错见 `references/conversion.md`。

## PPT 模板库（6 套 16:9）

`templates/ppt/` 提供 6 套成品模板，每套两种形态：`.pptx` 是 8 页示例 deck（直接改文字交付），`.potx` 是 0 页版式模板（供 `--template` 生成新 deck，不会带出示例页）。

| slug | 场景 | 适用 |
|:---|:---|:---|
| `business-report` | 商务办公 | 经营分析、季度汇报、项目复盘 |
| `consulting` | 商务办公 | 战略咨询、方案提案、路演 |
| `academic-defense` | 学术科研 | 学位论文答辩、课题结题 |
| `academic-seminar` | 学术科研 | 开题、组会、中期检查 |
| `government-report` | 政务公文 | 工作汇报、年度总结、督察整改 |
| `government-briefing` | 政务公文 | 政策解读、宣讲培训 |

```bash
# 套版式生成新 deck；.potx 无示例页，画幅随模板（16:9）
python -X utf8 scripts/create_powerpoint.py --output-path 答辩.pptx \
  --template templates/ppt/academic-defense.potx \
  --params '{"slides":[{"layout":"title","title":"论文题目","subtitle":"答辩人 · 导师 · 日期"}]}'
```

配色、页序、字体与每页用途见 `assets/manifest.json`；完整用法与约束见 `templates/README.md`；重新生成模板用 `python -X utf8 tools/build_ppt_templates.py`。

## 常用示例

```bash
# 1) 创建 Word：标题 + 段落 + 表格
python -X utf8 scripts/create_word_document.py --output-path 报告.docx --params '{
  "content": {"sections": [
    {"type": "heading", "text": "季度报告", "level": 1},
    {"type": "paragraph", "text": "本季度完成核心模块开发。"},
    {"type": "table", "table": {"headers": ["指标","值"], "rows": [["交付率","97%"]]}}
  ]},
  "scene": "report"
}'

# 2) 长 Markdown 转 Word
python -X utf8 scripts/create_word_from_markdown.py --source-path 长文.md --output-path 长文.docx --scene report

# 3) 读取并检查 Word
python -X utf8 scripts/read_word_document.py --file-path 报告.docx
python -X utf8 scripts/validate_word_document.py --file-path 报告.docx

# 4) 套模板填充占位符
python -X utf8 scripts/merge_office_template.py \
  --template-path templates/word_template.docx --output-path 输出.docx \
  --params '{"data":{"company":"示例公司","footer":"内部资料","title":"项目总结","summary":"摘要","body":"正文"}}'

# 5) PPT 配色/布局审计/清理/自动重排
python -X utf8 scripts/ppt_list_color_schemes.py
python -X utf8 scripts/ppt_apply_color_scheme.py --file-path 演示.pptx --scheme-name professional-blue
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path 演示.pptx   # 只检查越界/重叠
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path 演示.pptx --reflow  # 自动重排并写回
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path 演示.pptx --fix-spacing  # 旧的纵向间距优化
python -X utf8 scripts/ppt_clean_unused_placeholders.py --file-path 演示.pptx

# 6) 导出 PDF 并渲染检查（无 dsoffice 时自动回退 Office/WPS COM）
OFFICE_CONVERT_BACKEND=auto python -X utf8 scripts/convert_to_pdf.py --input-path 报告.docx --output-path 报告.pdf
python -X utf8 scripts/render_document_pages.py --input-path 报告.pdf --output-dir 报告_pages --pages 1,2
```

示例参数文件见 `assets/examples/word_report.json`、`assets/examples/excel_table.json`、`assets/examples/ppt_deck.json`、`assets/examples/framework_diagram.json`；用 `--params-file` 传入。模板说明见 `templates/README.md`。

## 任务→脚本组合

| 任务 | 推荐组合（可增减） |
|:---|:---|
| 从零写报告/公文/论文 | `create_word_document`（`scene` 选排版）→ `validate_word_document` → 需要时 `convert_to_pdf` + `render_document_pages` |
| 长 Markdown 转 Word | `create_word_from_markdown` → `read_word_document` 抽查 → `validate_word_document` |
| 改已有 Word | `read_word_document` 看结构 → `edit_word_document` 追加 → `validate_word_document` |
| 套模板/批量填充 | `ppt_analyze_template`（PPT）或 `read_word_document`（Word）看结构 → `merge_office_template` → `validate_word_document`/`read_*` 复核 |
| 做 Excel 报表 | `create_excel_workbook` → 需要公式结果时 `recalculate_excel_formulas` → `read_excel_workbook` |
| Excel 图表/条件格式/验证 | `create_excel_workbook`（`charts`/`conditional_formats`/`data_validations`）或 `excel_add_chart`（已有文件）→ `read_excel_workbook` |
| Word 批注/属性 | `word_manage_comments`（增/列）→ `word_set_properties`（读写属性） |
| PPT 已有文件加链接 | `read_powerpoint` 看形状 → `ppt_set_hyperlinks`（`shape_index`/`shape_text`） |
| 在打开着的 PPT 里加形状 | `ppt_manage_shapes`（`action=add`，`slide` + `shape`）；文件已打开走活动文档 COM，无需关闭；未打开时自动走文件级并备份 |
| 做 PPT | 按场景选 `templates/ppt/<slug>.potx`（见 `assets/manifest.json`）→ `create_powerpoint --template` → 需要时 `ppt_apply_color_scheme` → `ppt_optimize_slide_layout`（审阅 warnings；需要时 `--reflow` 或 `--fix-spacing`）→ `ppt_clean_unused_placeholders` → `render_document_pages` |
| PPT 框架图 | `ppt_create_framework_diagram`（文件级）或 `ppt_add_flowchart_slide`（活动文档） |
| 精细改某段/某格/某形状 | `inspect_office_document` 取路径 → `batch_edit_office_document`（原子）→ `read_*` 复核 |
| 转 PDF/预览 | `convert_to_pdf`（自动回退 COM）→ `render_document_pages` → 需要时用 pymupdf 读 span 字号（见 `references/render_qa.md`） |
| 活动文档实时改 | `get_open_office_documents` → `edit_open_word_document`/`edit_open_excel_workbook`/`edit_open_powerpoint` |
| 改已打开文档的某节/某段（不重生成） | `get_open_office_documents` → `replace_open_document_content`（`replace_sections` 按标题替换整节，`replace_texts` 查找替换）→ `read_word_document` 复核 |

## 输出与退出码

stdout 默认只输出一个 JSON 对象，不提供 `--json` 开关；诊断写 stderr；`--help` 是唯一非 JSON 输出。

成功：

```json
{"ok": true, "tool": "create_word_document", "data": {"success": true, "file_path": "报告.docx", "file_size": 12345, "message": "..."}}
```

失败：

```json
{"ok": false, "tool": "read_word_document", "error": "文件不存在: 报告.docx"}
```

| 退出码 | 含义 |
|:---|:---|
| `0` | 成功 |
| `1` | 门控不通过：`validate_word_document` 报告 `errors > 0` |
| `2` | 参数/输入错误：缺参数、JSON 非法、Pydantic 校验失败、枚举值非法 |
| `3` | 运行时/资产/外部调用错误：文件不存在、格式不支持、缺依赖、dsoffice/COM 失败、未预期异常 |

- 上游工具返回 `success: false` 时，CLI 转成 `ok: false` + 退出码 3，不把失败藏在 `data` 里。
- `read_word_document`、`read_excel_workbook`、`read_powerpoint`、`read_pdf_document`、`read_txt_document` 的 `data` 为 `{"markdown": "...", "result": {...}}`；`read_pdf_structured` 直接返回上游结构化 dict。
- `ppt_list_color_schemes` 的列表放在 `data.schemes`。
- `render_document_pages` 的 `page_count` 以输出目录实际 PNG 数为准；上游元数据不一致时返回 `page_count_reported`；`backend` 为 `dsoffice`/`pymupdf`/`com+pymupdf`。
- `convert_to_pdf` 返回 `backend_kind`（`dsoffice`/`com`）、`com_engine`、`com_prog_id`、`com_application`；回退时 `fallback_from[]` 记录上一后端与错误。`font_substitutions[]` 给出缺失等宽字体的本机可用回退。
- `recalculate_excel_formulas` 的 COM 回退返回 `calculation=CalculateFullRebuild|Calculate`；只支持 `.xlsx`。
- `batch_edit_office_document` 省略 `--output-path` 时默认生成 `*.bak-YYYYmmdd-HHMMSS` 并在结果返回 `backup_path`；`--no-backup` 关闭。
- `ppt_optimize_slide_layout` 默认只返回 `warnings[]`（`out_of_bounds`/`overlap`），不改文件；`--reflow` 写回并返回 `reflow[]`（`shape_id`/`action`/`reason`/`from`/`to`）与 `remaining_warnings[]`；`--fix-spacing` 只做旧的纵向间距优化。
- `validate_word_document` 的 `checks[]` 每项含稳定 `id`、`severity`、`message`、`details`；新检查均为 warning，`errors>0` 时退出码 1。
- 所有未预期异常由入口统一兜底为 JSON + 退出码 3，不会输出裸 traceback 到 stdout。

## 铁律

1. **路径来自参数**：脚本内不写死项目/用户目录；所有文档路径由调用方传入。解释器路径只出现在文档里。
2. **零 MCP 依赖**：不导入 `mcp`、不调用 MCP 工具、不启动 MCP 服务；`scripts/office_engine/` 是 vendored 纯 Python 引擎。
3. **不伪造结果**：不编造文档内容、数据、引用；`validate_word_document` 不通过就修，不假装通过；没渲染就不声称“版式已验证”。
4. **默认拒绝覆盖**：所有写文件的入口默认 `if_exists=error`，空白文件（不存在或 0 字节）可直接写；覆盖非空目标必须同时传 `--if-exists overwrite --confirm-overwrite`，否则退出码 2。原地写回前自动生成 `*.bak-YYYYmmdd-HHMMSS` 备份；`batch_edit_office_document` 省略 `output_path` 时也会先备份，`--no-backup` 可关闭，关闭后需自行确认旧内容可丢弃。不要用文件系统命令绕过闸门。
5. **失败要可见**：工具 `success=false` 必须是 `ok=false` + 退出码 3；参数错误必须是退出码 2，不混用。
6. **活动文档不启动应用**：活动文档 COM 入口只连接已打开窗口，找不到就报错。转换/渲染/重算的 COM 回退是受控例外：`DispatchEx` 独立隐藏实例，完成或失败都 Quit，不触碰已打开文档。
7. **模板/示例可分发**：`templates/` 与 `assets/examples/` 只含自制合成内容；客户文档、未公开数据、第三方全文不放进 skill。
8. **渲染验收才算版式验收**：结构检查与 reflow 通过不等于视觉正确；需要看版式时走 `convert_to_pdf` + `render_document_pages`，或人工打开确认。
9. **回退要可见**：`fallback_from[]`、`backend_kind`、`com_engine` 必须如实返回；不要因为回退成功就隐藏 dsoffice 缺失或 COM 失败原因。
10. **禁止复写已有二进制文件**：.doc/.docx/.ppt/.pptx/.xls/.xlsx 等已有文件只能原地修改（活动文档 COM 或文件级编辑入口）；`create_*` 重生成入口只允许用于新建文件、空白文件，或用户明确授意的整体重排。修改后核对关键对象仍在（如 OLE 数量、修订/批注），不把「保存成功」当作「内容未被破坏」。

## references 索引

| 遇到什么 | 读哪个文件 |
|:---|:---|
| WPS/Office 活动文档、COM 转换回退、超时/只读/进程清理、找不到窗口 | `references/wps_com.md` |
| 中文字号、场景排版、页眉页脚/页码域、PDF 字体 | `references/fonts_page_numbers.md` |
| Word/Excel 表格样式、PPT 布局/配色/框架图、`--reflow` | `references/tables_slides.md` |
| 渲染预览、pymupdf 读 span 字号、缺字体、视觉验收 | `references/render_qa.md` |
| 选模板/查素材/模板用法 | `assets/manifest.json`、`templates/README.md` |
| 45 个入口的 flag、`--params` 字段、JSON 示例 | `references/cli_args.md` |
| Word/Excel/PPT 逐项功能面与覆盖状态 | `references/capabilities_word.md`、`capabilities_excel.md`、`capabilities_ppt.md` |
| C# office-native 构建、命令、COM 边界 | `tools/office-native/README.md` |
| dsoffice/Office/WPS 后端顺序、转换/渲染/重算回退、OCR | `references/conversion.md` |

工程文件：`README.md`（安装/运行时/测试/边界）、`config/defaults.json`（超时与渲染默认值）、`templates/`（起步模板 + 6 套 16:9 PPT 模板）、`assets/`（素材清单 `manifest.json`、示例与缓存说明）、`tests/`（契约/结构/触发/parity）。

## 验证状态

- **2026-11-20 PPT 模板库实测**：`templates/ppt/` 6 套 16:9 模板（每套 `.pptx` 8 页示例 + `.potx` 0 页版式）；用 `academic-defense.potx` 经 `create_powerpoint --template` 生成 2 页，画幅 13.33×7.5、`Title Slide`/`Title and Content` 索引正确，标题 30pt 主题色左对齐、正文 18pt 由版式 `lstStyle` 继承；6 套封面/内容/图表/表格/两栏页经 PDF→PNG 渲染核对。
- **自动化**：2026-11-20 引擎切换后全量 **64 项通过、0 跳过、0 失败**；复跑 `python -X utf8 -m unittest discover -s tests -v`。覆盖 45 个入口 `--help`、固定输出全量比对、退出码 1/2/3、创建→读取→校验往返、结构/触发基线、字体回退、渲染计数修正、批量备份、PPT 越界/重叠审计与 `--reflow`、dsoffice 与真实 Office/WPS COM、COM 缺失回退、17 项 validate fixture、扫描件 OCR，以及覆盖闸门（拒绝非空、空白放行、确认覆盖生成备份、编辑生成备份）。
- **2026-11-20 `ppt_manage_shapes` 实测**：活动文档 add 返回 `mode=open_document`（文件开着时直接加到当前窗口并保存，无需关闭）；文件级 add 返回 `mode=file` 并生成 `.bak-*`；list 两条路径的索引/名称/位置/文本一致（`type` 在文件级为 XML 元素名、在活动文档为 Mso 类型码；`preset` 仅文件级输出）；越界 slide 返回退出码 2；渲染 PNG 核对形状位置与文字正确。
- **2026-11-20 引擎策略调整实测**：默认改为 Python 优先，`ppt_manage_shapes` 同时具备 Python（活动文档 COM + python-pptx 文件级）与 C# 实现；三路径（默认 Python 文件级、显式 C#、活动文档 COM）add/list 与越界/非法颜色/非法尺寸退出码均通过；全量 64 项通过、0 跳过；7 个 C# 专属入口仍自动走 C#。
- **2026-11-20 C# 视觉对齐实测**：`create_powerpoint` 的表格样式（内置 Medium Style 2 - Accent 1）、正文占位符空段与列表顺序、形状 `anchor="ctr"`、图表标题/轴/图例字号修复后，同定义双引擎内容页像素差异从 15.835% 降至 4.391%；结构差异仅剩列表承载方式（C# 占位符 vs Python 文本框）；全量 64 项通过。对比方法与数据见项目 `docs/06-PPT双引擎效果对比-2026-11.md`。
- **2026-11-19 扩展实测**：Word 域（PAGE/NUMPAGES/DATE）、超链接、批注增/列、核心属性、TOC/分节追加；Excel 8 类图表（含散点 X/Y 引用）、条件格式、数据验证、超级表总计行、定义名称、批注、图片、页面设置；PPT 表格/图表几何、图片裁剪、形状/文字超链接、真项目符号/编号、双栏占位符、`.potx` 模板。全量 45/47 在当前非提权会话通过；2 项依赖 fake dsoffice `.cmd` 的测试在提权会话 2/2 通过。
- **2026-11-19 office-native 实测**：Word 修订插入/接受、删除/拒绝、脚注、内容控件（文本/日期/下拉/复选框）增删改查；PPT 动画/切换/批注/SmartArt（134 个布局）；Excel 透视表创建/列出/刷新与 PDF 导出；PowerPoint/Excel 进程无残留。
- **已实测（2026-11-15）**：dsoffice 导出 docx/xlsx/pptx→PDF 均成功；Word/Excel/PowerPoint COM 与 WPS `KWps/KET/KWPP` 均可导出；`--pages 1 --dpi 144` 渲染 1191×1684 PNG；公式重算回读合计 255/390；validate 干净 fixture 17/17、问题 fixture 命中 8 类 warning；reflow 四类形状重排后 `remaining_warning_count=0`。
- **未验证 / 边界**：COM 回退的 Excel 渲染按 sheet + print area 近似 `--cell-range`，复杂分页/缩放可能与 dsoffice 不同；不做逐像素视觉验收；仅在 Windows + Office/WPS 环境验证，其他平台未覆盖。
- **复跑**：`python -X utf8 -m unittest discover -s tests -v`。
