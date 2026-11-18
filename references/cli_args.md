# CLI 参数参考

来源：本 skill 的 32 个入口脚本（mcp-office-writer 0.6.0 的 1:1 包装，2026-11）。

适用边界：命令行调用与 `--params` JSON 的字段名；Pydantic 模型的完整字段见 `scripts/office_engine/models.py`。所有入口都支持 `--params '<JSON 对象>'` 与 `--params-file <路径>`，同名显式 flag 优先；`--params-file -` 从 stdin 读取。

## 通用规则

- 默认 stdout 只输出一个 JSON 对象；`--help` 是唯一例外。
- 复杂参数（`content`、`sections`、`sheets`、`slides`、`diagram`、`operations`、`append_*`、`data`）放在 `--params` 里，由对应 Pydantic 模型 `model_validate` 校验。
- 简单参数优先用 flag：`--output-path`、`--file-path`、`--input-path`、`--if-exists`、`--max-chars` 等。
- `if_exists`：`overwrite`（覆盖）/ `error`（报错）/ `rename`（自动改名）；`render_document_pages` 只支持 `overwrite`/`error`。
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

`batch_edit_office_document` 省略 `--output-path` 时默认先在同目录写 `*.bak-YYYYmmdd-HHMMSS`，结果带 `backup_path`；`--no-backup` 关闭备份。

`SectionItem.type`：`heading` / `paragraph` / `table` / `image` / `page_break` / `bullet_list` / `numbered_list` / `code` / `toc` / `section_break`。

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

`validate_word_document` 的 `data.checks[]` 每项含稳定 `id`：`file_size`、`paragraph_count`、`blank_pages`、`line_spacing`、`table_integrity`、`heading_levels`、`images`、`sections`、`header_footer`、`placeholders`、`font_installed`、`heading_numbering`、`page_number_field`、`toc_field`、`table_columns`、`empty_paragraph_ratio`、`mixed_fonts`；`severity` 为 `info/warn/error`，`errors>0` 时退出码 1。

`read_word/excel/powerpoint/pdf/txt` 的 `data` 是 `{"markdown": "...", "result": {...}}`：`markdown` 与上游 FastMCP 的文本返回一致，`result` 是上游扁平 dict。`read_pdf_structured` 直接返回上游结构化 dict。

## 转换 / 渲染 / 重算

| 入口 | 关键 flag |
|:---|:---|
| `convert_to_pdf` | `--input-path` `--output-path` `--sheet` `--if-exists` |
| `render_document_pages` | `--input-path` `--output-dir` `--pages` `--dpi` `--sheet` `--cell-range` `--if-exists` |
| `recalculate_excel_formulas` | `--input-path` `--output-path` `--if-exists` |

`--pages all|1,3,5`；`--cell-range` 必须与 `--sheet` 同时使用；`--dpi` 默认 144，上游夹到 72–600。`render_document_pages` 的 `page_count` 按实际 PNG 产物修正，上游值不一致时放在 `page_count_reported`。

后端由 `OFFICE_CONVERT_BACKEND=auto|dsoffice|com|none` 控制（默认 auto：dsoffice → Microsoft Office COM → WPS COM）。`convert_to_pdf` 返回 `backend`、`backend_kind`、`com_engine`、`com_prog_id`、`com_application`、`font_substitutions[]`；回退时返回 `fallback_from[{backend,error}]`。`render_document_pages` 返回 `backend=dsoffice|pymupdf|com+pymupdf`。`recalculate_excel_formulas` 的 COM 回退返回 `calculation=CalculateFullRebuild|Calculate`，只支持 `.xlsx`。`MCP_OFFICE_DSOFFICE` 指定 dsoffice 路径；`MODULE_OFFICE_DSOFFICE` 是兼容别名。

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

`ppt_optimize_slide_layout` 默认只审计并返回 `warnings[]`；`--reflow` 写回并返回 `reflow[]`（`shape_id`、`action=move/resize/move_resize`、`reason`、`from`/`to` 英寸坐标）、`remaining_warnings[]`、`remaining_warning_count`；`--fix-spacing` 走旧的纵向间距优化。`--reflow` 优先于 `--fix-spacing`。

## 活动文档（Windows COM）

| 入口 | 关键 flag / params |
|:---|:---|
| `get_open_office_documents` | 无 |
| `edit_open_word_document` | `--file-path`；`append_sections`、`default_font` |
| `edit_open_excel_workbook` | `--file-path`；`append_sheets` |
| `edit_open_powerpoint` | `--file-path`；`append_slides` |

这些入口只连接已打开的应用，不启动 Word/WPS/Excel/PowerPoint；60 秒超时返回失败。

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
