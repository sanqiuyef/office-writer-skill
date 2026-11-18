---
name: office-writer
description: 用 32 个独立 CLI 入口创建、编辑、读取、转换和验收 Word/Excel/PPT/PDF：写报告/公文/论文/简历/合同、Markdown 转 Word、公式重算、PPT 配色/布局/清理/框架图、套模板填充 {{占位符}}、导出 PDF、渲染页面、检查字号字体、批量编辑与活动 Office 文档。转换/渲染/重算支持 dsoffice → Microsoft Office COM → WPS COM 回退。触发：写 Word、生成 docx、写报告、写公文、写论文、写简历、写合同、改稿、追加段落、写 Excel、生成表格、公式重算、写 PPT、做演示文稿、PPT 美化、PPT 配色、PPT 优化布局、清理占位符、生成框架图、写 PDF、Markdown 转 Word、套模板、填充占位符、导出 PDF、转 PDF、渲染页面、检查排版、检查字号、读取文档、批量编辑文档、活动 Office 文档。
---

# Office Writer

一个能力一个入口，32 个脚本与 mcp-office-writer 0.6.0 的 32 个工具 1:1 对应。所有文档逻辑在 `scripts/office_engine/`（vendored），本 skill 只做 CLI 适配、JSON 契约和验收辅助，不启动 MCP 服务。

## 触发条件

用户要求创建、修改、读取、转换或验收办公文档时使用：

- Word：写报告/公文/论文/简历/合同、套模板、Markdown 转 Word、追加内容、目录、页眉页脚、封面。
- Excel：多工作表、表头/格式/合并/冻结/筛选、公式写入与重算。
- PowerPoint：新建演示文稿、追加幻灯片、配色、布局优化、清理占位符、分析模板、生成框架图/流程图。
- PDF：新建 PDF、Office 转 PDF、逐页渲染 PNG、结构化读取、字号/字体验收。
- 结构化编辑：按 `/body/p[1]`、`/Sheet1/A1`、`/slide[1]/shape[2]` 路径原子批量编辑。
- 活动文档：Word/WPS、Excel/WPS 表格、PowerPoint 已打开文档的实时追加。

## 何时不要用

- 只想“想文案/改措辞”，不需要生成或修改文档 → 直接写文本或使用写作类 skill。
- 只做 Excel 数据排序、透视、统计建模 → 用 pandas/openpyxl 脚本；本 skill 负责版式与写入。
- 只做 PDF 拆分/合并/加密/OCR 校对 → 用专门的 PDF 工具；本 skill 的 OCR 仅用于读取扫描件。
- 需要生成示意图/海报/插画 → 用绘图或生图工具；`ppt_create_framework_diagram` 只画阶段+节点+箭头。
- 需要启动 Word/WPS/Excel/PowerPoint → 活动文档入口只连接现有窗口，不启动应用；转换/渲染/重算的 COM 回退是例外：它用 `DispatchEx` 启动隐藏独立实例，完成后 Quit。
- 非 Windows 环境需要 COM 活动文档或 COM 转换 → 用文件级入口并安装 dsoffice；COM 回退会明确报不可用。

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

参数细节与 JSON 模型见 `references/cli_args.md`；完整模型定义在 `scripts/office_engine/models.py`。转换/渲染/重算的后端顺序用 `OFFICE_CONVERT_BACKEND=auto|dsoffice|com|none` 控制：默认 auto 依次尝试 dsoffice、Microsoft Office COM、WPS COM；dsoffice 安装与 COM 排错见 `references/conversion.md`。

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
| 做 PPT | `ppt_list_color_schemes` 选配色 → `create_powerpoint` → `ppt_apply_color_scheme` → `ppt_optimize_slide_layout`（审阅 warnings；需要时 `--reflow` 或 `--fix-spacing`）→ `ppt_clean_unused_placeholders` → `render_document_pages` |
| PPT 框架图 | `ppt_create_framework_diagram`（文件级）或 `ppt_add_flowchart_slide`（活动文档） |
| 精细改某段/某格/某形状 | `inspect_office_document` 取路径 → `batch_edit_office_document`（原子）→ `read_*` 复核 |
| 转 PDF/预览 | `convert_to_pdf`（自动回退 COM）→ `render_document_pages` → 需要时用 pymupdf 读 span 字号（见 `references/render_qa.md`） |
| 活动文档实时改 | `get_open_office_documents` → `edit_open_word_document`/`edit_open_excel_workbook`/`edit_open_powerpoint` |

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
4. **不静默覆盖**：默认 `if_exists=overwrite`，但覆盖前确认目标；需要保留旧文件时用 `rename` 或显式备份。`batch_edit_office_document` 省略 `output_path` 会自动先写 `*.bak-YYYYmmdd-HHMMSS`，`--no-backup` 可关闭；关闭后需自行确认旧内容可丢弃。
5. **失败要可见**：工具 `success=false` 必须是 `ok=false` + 退出码 3；参数错误必须是退出码 2，不混用。
6. **活动文档不启动应用**：活动文档 COM 入口只连接已打开窗口，找不到就报错。转换/渲染/重算的 COM 回退是受控例外：`DispatchEx` 独立隐藏实例，完成或失败都 Quit，不触碰已打开文档。
7. **模板/示例可分发**：`templates/` 与 `assets/examples/` 只含自制合成内容；客户文档、未公开数据、第三方全文不放进 skill。
8. **渲染验收才算版式验收**：结构检查与 reflow 通过不等于视觉正确；需要看版式时走 `convert_to_pdf` + `render_document_pages`，或人工打开确认。
9. **回退要可见**：`fallback_from[]`、`backend_kind`、`com_engine` 必须如实返回；不要因为回退成功就隐藏 dsoffice 缺失或 COM 失败原因。

## references 索引

| 遇到什么 | 读哪个文件 |
|:---|:---|
| WPS/Office 活动文档、COM 转换回退、超时/只读/进程清理、找不到窗口 | `references/wps_com.md` |
| 中文字号、场景排版、页眉页脚/页码域、PDF 字体 | `references/fonts_page_numbers.md` |
| Word/Excel 表格样式、PPT 布局/配色/框架图、`--reflow` | `references/tables_slides.md` |
| 渲染预览、pymupdf 读 span 字号、缺字体、视觉验收 | `references/render_qa.md` |
| 32 个入口的 flag、`--params` 字段、JSON 示例 | `references/cli_args.md` |
| dsoffice/Office/WPS 后端顺序、转换/渲染/重算回退、OCR | `references/conversion.md` |

工程文件：`README.md`（安装/运行时/测试/边界）、`config/defaults.json`（超时与渲染默认值）、`templates/`（起步模板）、`assets/`（示例与缓存说明）、`tests/`（契约/结构/触发/parity）。
