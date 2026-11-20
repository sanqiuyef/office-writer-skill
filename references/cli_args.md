# CLI 参数参考

来源：本 skill 的 44 个入口脚本（32 个为 mcp-office-writer 0.6.0 的 1:1 包装 + 12 个本地入口：活动文档替换、Word 批注/属性、Excel 图表、PPT 超链接，以及 7 个 office-native 包装，2026-11）。

适用边界：命令行调用与 `--params` JSON 的字段名；全部模型的字段类型、默认值与说明见 `references/models_word.md`、`models_excel.md`、`models_ppt.md`、`models_pdf_common.md`，模型定义源为 `scripts/office_engine/models.py`。所有入口都支持 `--params '<JSON 对象>'` 与 `--params-file <路径>`，同名显式 flag 优先；`--params-file -` 从 stdin 读取。

## 通用规则

- 默认 stdout 只输出一个 JSON 对象；`--help` 是唯一例外。
- 复杂参数（`content`、`sheets`、`slides`、`diagram`、`operations`、`chart`、`links`、`properties`、`append_*` 等）放在 `--params` 里，由对应 Pydantic 模型 `model_validate` 校验；`merge_office_template` 的 `data` 是自由对象（只检查为 dict），不套固定模型。
- 简单参数优先用 flag：`--output-path`、`--file-path`、`--input-path`、`--if-exists`、`--max-chars` 等。`--reflow`、`--fix-spacing`、`--no-backup` 也可在 `--params` 里传同名布尔键（显式 flag 优先）。
- `if_exists`：`error`（默认，拒绝覆盖非空目标）/ `rename`（自动改名）/ `overwrite`（覆盖，必须同时加 `--confirm-overwrite`）；`render_document_pages` 只支持 `error`/`overwrite`。空白文件（不存在或 0 字节）和空目录不受此限。覆盖发生时自动生成 `*.bak-YYYYmmdd-HHMMSS` 备份；目录目标整体改名为 `.bak-*`，不删除旧内容。该策略只作用于有输出目标的入口；`edit_*`、`batch_edit_office_document` 与 PPT/Excel 增强入口原位写回并自动备份。
- 路径建议使用绝对路径；相对路径按调用时的工作目录解析。

## 创建 / 编辑

| 入口 | 关键 flag | 关键 params |
|:---|:---|:---|
| `create_word_document` | `--output-path` `--template` `--scene` `--if-exists` | `content: WordContent`、`page_setup`、`header_footer`、`default_font`、`cover_config` |
| `edit_word_document` | `--file-path` | `append_sections: SectionItem[]`、`default_font` |
| `create_word_from_markdown` | `--source-path` `--output-path` `--template` `--scene` `--if-exists` | 无复杂模型；source 必须存在 |
| `create_excel_workbook` | `--output-path` `--template` `--if-exists` | `sheets: SheetDef[]` |
| `edit_excel_workbook` | `--file-path` | `append_sheets: SheetDef[]` |
| `create_powerpoint` | `--output-path` `--template` `--if-exists` | `slides: SlideDef[]` |
| `edit_powerpoint` | `--file-path` | `append_slides: SlideDef[]` |
| `create_pdf_document` | `--output-path` `--if-exists` | `content: PdfContent`、`page_setup`、`font_config` |
| `merge_office_template` | `--template-path` `--output-path` `--if-exists` | `data: object`（支持 `{{user.name}}`、`{{items[0].title}}`） |
| `batch_edit_office_document` | `--file-path` `--output-path` `--no-backup` | `operations: OfficeOperation[]` |
| `word_manage_comments` | `--file-path` `--list` | `add: WordCommentDef[]`（`text`/`author`/`initials`/`anchor_text`/`paragraph_index`/`match_case`）；`list: bool` |
| `word_set_properties` | `--file-path` `--list` | `properties: WordProperties`（title/subject/author/keywords/comments/category/last_modified_by/created/modified） |
| `excel_add_chart` | `--file-path` `--sheet` | `chart: ExcelChartDef` |
| `ppt_set_hyperlinks` | `--file-path` | `links: PptHyperlinkDef[]`（`slide`/`url`/`shape_index`/`shape_text`/`text_match`/`match_case`） |

`batch_edit_office_document` 省略 `--output-path` 时默认先在同目录写 `*.bak-YYYYmmdd-HHMMSS`，结果带 `backup_path`；`--no-backup` 关闭备份。

`SectionItem.type`：`heading` / `paragraph` / `table` / `image` / `page_break` / `bullet_list` / `numbered_list` / `code` / `toc` / `section_break` / `hyperlink` / `field` / `bookmark`。`hyperlink` 用 `{"text","url","tooltip"}`；`field` 用 `{"field_type":"PAGE|NUMPAGES|SECTIONPAGES|DATE|TIME|FILENAME|AUTHOR|TITLE|CUSTOM","code","fallback_text","number_format"}`，CUSTOM 时 `code` 必填；`bookmark` 用 `{"name"}`，`text` 为书签内容。

`WordContent.properties` 可在创建时写核心属性；`PageSetup` 另支持 `columns`（1-4）、`column_spacing`、`page_border`（single/double/dashed/dotted）、`page_border_color`、`watermark_text`、`watermark_color`。`SheetDef` 的扩展字段：`row_heights`、`sheet_state`、`charts`、`conditional_formats`、`data_validations`、`tables`、`defined_names`、`hyperlinks`、`comments`、`images`、`page_setup`、`protection`、`sort`。PPT 的 `SlideTableDef`/`ChartDef` 支持 `x/y/width/height`，`SlideImageDef` 支持 `crop_left/right/top/bottom` 与 `hyperlink_url`，`SlideShapeDef` 支持 `hyperlink_url`；`SlideDef` 支持 `footer_text`/`show_slide_number`，`PptContent` 支持 `sections`/`show_settings`（仅创建）。

`OfficeOperation`：`{"op":"set|add|remove|move|replace","path":"/body/p[1]","element_type":"paragraph","props":{...},"target_path":"...","position":"before|after"}`。

## 读取

| 入口 | 关键 flag |
|:---|:---|
| `read_word_document` | `--file-path` `--max-chars` |
| `read_excel_workbook` | `--file-path` `--max-rows-per-sheet` `--max-chars` |
| `read_powerpoint` | `--file-path` `--max-chars` |
| `read_pdf_document` | `--file-path` `--max-chars` |
| `read_pdf_structured` | `--file-path` `--max-chars` `--include-blocks` `--extract-images` |
| `read_txt_document` | `--file-path` `--max-chars` |
| `inspect_office_document` | `--file-path` `--max-nodes`（1–10000） |
| `validate_word_document` | `--file-path` |

读取类默认值：`--max-chars=50000`（`read_word/excel/powerpoint/pdf/txt` 与 `read_pdf_structured`）、`--max-rows-per-sheet=1000`（Excel）、`--max-nodes=1000`（inspect，可到 10000）。`read_excel_workbook` 按 `cellXfs → numFmtId → numFmts` 识别日期/时间/时长并输出 datetime 字符串；C# 后端按行流式读取，达到上限立即停止。

`validate_word_document` 的 `data.checks[]` 每项含稳定 `id`：`file_size`、`paragraph_count`、`blank_pages`、`line_spacing`、`table_integrity`、`heading_levels`、`images`、`sections`、`header_footer`、`placeholders`、`font_installed`、`heading_numbering`、`page_number_field`、`toc_field`、`table_columns`、`empty_paragraph_ratio`、`mixed_fonts`；`severity` 为 `info/warn/error`，`errors>0` 时退出码 1。

`read_word/excel/powerpoint/pdf/txt` 的 `data` 是 `{"markdown": "...", "result": {...}}`：`markdown` 与上游 FastMCP 的文本返回一致，`result` 是上游扁平 dict。`read_pdf_structured` 直接返回上游结构化 dict。

## 转换 / 渲染 / 重算

| 入口 | 关键 flag |
|:---|:---|
| `convert_to_pdf` | `--input-path` `--output-path` `--sheet` `--if-exists` |
| `render_document_pages` | `--input-path` `--output-dir` `--pages` `--dpi` `--sheet` `--cell-range` `--if-exists` |
| `recalculate_excel_formulas` | `--input-path` `--output-path` `--if-exists` |

`--pages all|1,3,5`；`--cell-range` 必须与 `--sheet` 同时使用（Excel 输入强制同时给）；`--dpi` 默认 144，上游夹到 72–600。默认输出：`convert_to_pdf` 为 `<input>.pdf`、`render_document_pages` 为 `<input>_pages`、`recalculate_excel_formulas` 为 `<stem>_recalc<suffix>`；`create_word_from_markdown` 的 `--scene` 默认 `report`。`render_document_pages` 的 `page_count` 按实际 PNG 产物修正，上游值不一致时放在 `page_count_reported`。

后端由 `OFFICE_CONVERT_BACKEND=auto|dsoffice|com|none` 控制（默认 auto：dsoffice → Microsoft Office COM → WPS COM）。`convert_to_pdf` 在 dsoffice 分支返回 `backend`、`backend_kind=dsoffice`、`missing_fonts[]`、`font_substitutions[]`；COM 分支才返回 `com_engine`、`com_prog_id`、`com_application`。回退时返回 `fallback_from[{backend,error}]`。`render_document_pages` 的 `backend` 取值：C# 默认 `pdfium`（PDF 直接渲染）/`dsoffice+pdfium`/`com+pdfium`，Python 回退引擎为 `dsoffice`/`pymupdf`/`com+pymupdf`。`recalculate_excel_formulas` 的 COM 回退返回 `calculation=CalculateFullRebuild|Calculate`，只支持 `.xlsx`。`MCP_OFFICE_DSOFFICE` 指定 dsoffice 路径；`MODULE_OFFICE_DSOFFICE` 是兼容别名。

## PPT 增强

| 入口 | 关键 flag / params |
|:---|:---|
| `ppt_list_color_schemes` | 无 |
| `ppt_apply_color_scheme` | `--file-path` `--scheme-name`（默认 `professional-blue`） |
| `ppt_optimize_slide_layout` | `--file-path` `--reflow` `--fix-spacing` |
| `ppt_clean_unused_placeholders` | `--file-path` |
| `ppt_analyze_template` | `--file-path` |
| `ppt_create_framework_diagram` | `--output-path` `--if-exists`；`diagram: FrameworkDiagram` |
| `ppt_add_flowchart_slide` | `--file-path`；`diagram: FrameworkDiagram`（活动 PowerPoint） |
| `ppt_set_hyperlinks` | `--file-path`；`links: PptHyperlinkDef[]`（`slide` 1 起、`shape_index` 1 起或 `shape_text`、可选 `text_match`） |

`ppt_optimize_slide_layout` 默认只审计并返回 `warnings[]`；`--reflow` 写回并返回 `reflow[]`（`shape_id`、`action=move/resize/move_resize`、`reason`、`from`/`to` 英寸坐标）、`remaining_warnings[]`、`remaining_warning_count`；`--fix-spacing` 走旧的纵向间距优化。`--reflow` 优先于 `--fix-spacing`。

## 活动文档（Windows COM）

| 入口 | 关键 flag / params |
|:---|:---|
| `get_open_office_documents` | 无 |
| `replace_open_document_content` | `--file-path`；`replace_sections`（按标题替换整节）、`replace_texts`（查找替换） |
| `edit_open_word_document` | `--file-path`；`append_sections`、`default_font` |
| `edit_open_excel_workbook` | `--file-path`；`append_sheets` |
| `edit_open_powerpoint` | `--file-path`；`append_slides` |

这些入口只连接已打开的应用，不启动 Word/WPS/Excel/PowerPoint；60 秒超时返回失败。活动文档追加只实现基础字段：Excel 的图表/条件格式/数据验证/超级表/名称/批注/图片/页面设置/行高/隐藏等扩展字段、PPT 的表格几何/图片裁剪/超链接字段会明确报错（退出码 3），请关闭文件后改用文件级入口。`replace_open_document_content` 是本地新增入口（上游 0.6.0 无对应工具），用于"修改优先于重生成"：Word 按标题替换整节或全文查找替换，PPT/Excel 做查找替换，由应用自身写回以保留公式、修订、批注与嵌入对象。

## office-native 扩展

全部 44 个入口由 `tools/office-native` C# 实现（.NET 8 + Open XML SDK / Office COM），默认走 C#；`OFFICE_NATIVE_ENGINE=python` 可回退旧引擎。构建：`pwsh -File tools/office-native/build.ps1`（含 `IncludeNativeLibrariesForSelfExtract=true`，PDFium 原生库才会随单文件发布），或用 `OFFICE_NATIVE_BIN` 指向 `office-native.exe`。未构建时默认调用返回退出码 3 并提示构建命令。所有修改型命令写回前自动 `.bak-*`。

| 入口 | 关键 flag | 关键 params |
|:---|:---|:---|
| `word_manage_revisions` | `--file-path` `--action` | `action: list/accept-all/reject-all/track-insert/track-delete`；`text`、`anchor_text`、`paragraph_index`、`author`、`date` |
| `word_manage_footnotes` | `--file-path` `--action` | `action: add/list`；`text`、`anchor_text`、`paragraph_index` |
| `word_manage_content_controls` | `--file-path` `--action` `--tag` | `action: add-text/add-date/add-dropdown/add-checkbox/list/set/remove`；`tag`、`alias`、`text`、`date`、`date_format`、`items`、`value`、`checked` |
| `ppt_manage_effects` | `--file-path` `--action` `--slide` | `action: animations-add/list/clear、transition-set/list/clear、comments-add/list/delete`；`shape_index`、`effect`（名称或 `entry_effect` 数值）、`duration`、`advance_time`、`text`、`author` |
| `ppt_manage_smartart` | `--file-path` `--action` `--slide` | `action: layouts/add`；`layout`、`texts`、`x/y/width/height`（英寸） |
| `excel_manage_pivots` | `--file-path` `--action` `--sheet` | `action: create/list/refresh`；`source_sheet`、`source_range`、`target_sheet`、`target_cell`、`name`、`row_field`、`column_field`、`data_field`、`function` |
| `excel_print` | `--file-path` `--action` `--sheet` `--output` | `action: pdf/print`；`output`、`printer`、`copies`、`confirm_print`（print 必填） |

## 最小示例

```bash
# Word：标题 + 段落 + 表格
python -X utf8 scripts/create_word_document.py --output-path 报告.docx --params '{
  "content": {"sections": [
    {"type": "heading", "text": "季度报告", "level": 1},
    {"type": "paragraph", "text": "本季度完成核心模块开发。"},
    {"type": "table", "table": {"headers": ["指标","值"], "rows": [["交付率","97%"]]}}
  ]},
  "scene": "report"
}'

# Excel：一个工作表
python -X utf8 scripts/create_excel_workbook.py --output-path 数据.xlsx --params '{
  "sheets": [{"name":"数据","headers":["区域","数量"],"data":[["华北",120]],"freeze":"A2"}]
}'

# 读取：返回 markdown + result
python -X utf8 scripts/read_word_document.py --file-path 报告.docx

# 模板合并：data 支持嵌套对象与数组索引
python -X utf8 scripts/merge_office_template.py \
  --template-path templates/word_template.docx --output-path 合并.docx \
  --params '{"data":{"company":"示例公司","title":"项目总结","summary":"摘要","body":"正文","footer":"内部资料"}}'
```

## 错误定位

- 退出码 2：参数缺失/JSON 非法/Pydantic 校验失败/枚举值非法；
- 退出码 3：文件不存在、格式不支持、文档损坏、缺依赖、dsoffice/COM 失败；
- 错误信息在 `error` 字段；上游工具返回的 `success=false` 也会转成退出码 3。
