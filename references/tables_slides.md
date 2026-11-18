# 表格、幻灯片与框架图样式

来源：mcp-office-writer 0.6.0 的 `word_writer.py`、`excel_writer.py`、`ppt_writer.py`、`models.py`（2026-11）。

适用边界：用本 skill 的数据模型生成 Office 表格/幻灯片；不覆盖手工拖拽排版、SmartArt、复杂动画。

## Word 表格

`SectionItem(type="table", table={...})`：

```json
{
  "headers": ["指标", "目标", "实际"],
  "rows": [["交付率", "95%", "97%"], ["缺陷率", "2%", "1.4%"]],
  "column_widths": [1.8, 1.2, 1.2],
  "header_style": {"bold": true, "bg_color": "D9EAF7", "font_color": "000000"},
  "merged_cells": [{"range": "A1:B1"}]
}
```

- 表格样式固定为 `Table Grid`；表头样式通过 `header_style` 设置。
- `merged_cells` 使用 A1 风格范围；解析失败只记 warning，不中断生成。
- 列宽单位英寸；行/列为空时上游跳过并记 warning。
- 长文本没有自动换行控制；需要精确版式时用 `batch_edit_office_document.py` 设置段落属性，或生成后人工调整。

## Excel 工作簿

`SheetDef` 关键字段：

| 字段 | 说明 |
|:---|:---|
| `name` | 工作表名，超过 31 字符会被截断 |
| `headers` / `data` | 表头与数据；`data` 支持 str/int/float/None |
| `allow_formulas` | 默认 false；为 true 时以 `=` 开头的字符串写成公式 |
| `column_widths` | 字符数级列宽 |
| `header_style` / `column_formats` | 表头与按列格式 |
| `merged_cells` | A1 范围 |
| `freeze` | 冻结窗格，如 `A2` |
| `auto_filter` | 自动筛选 |
| `tab_color` | 标签颜色 HEX |

注意：openpyxl 写入的公式没有缓存值；`read_excel_workbook.py` 用 `data_only=True` 读取时看到空值。需要缓存值时跑 `recalculate_excel_formulas.py`（dsoffice 或 Excel/WPS COM 回退）。

## PPT 布局

`SlideDef.layout` 支持：

`title` / `section` / `content` / `two_content` / `comparison` / `title_only` / `blank` / `content_caption` / `picture_caption`

`body`、`left`、`right` 是 `SlideItem` 列表，元素类型：

`paragraph` / `table` / `image` / `shape` / `connector` / `bullet_list` / `numbered_list` / `chart`

- 图片 `SlideImageDef` 必须且只能提供 `path` 或 `url` 之一；URL 仅 http/https，单张 20MB 上限，缓存到 `assets/image_cache/`。
- 形状支持 `rect/rounded_rect/oval/arrow/cylinder/diamond`，可设置填充、边框、虚线、竖排文字。
- 图表类型 `column/bar/line/pie`，需要 `categories` 与 `series`。
- `notes` 写演讲者备注；`bg_color` 设幻灯片背景；`slide_width/slide_height` 单位英寸。

## 配色方案与美化

- `ppt_list_color_schemes.py` 返回 12 套方案：商务蓝、暗夜模式、暖色大地、极简优雅、活力缤纷、企业绿、科技紫、经典中国红、清新自然、海洋蓝、日落暖橙、高级灰。
- `ppt_apply_color_scheme.py --scheme-name professional-blue` 修改标题、正文、背景、形状填充、表格表头。
- `ppt_optimize_slide_layout.py` 默认只读审计越界/重叠；`--fix-spacing` 按 Y 坐标均匀分布带文字形状，并按文本行数调整高度；`--reflow` 自动重排越界/重叠形状并写回（0.1in 页边距、0.15in 最小间距，必要时按比例缩小），只改几何，不碰字体/颜色/文本。
- `--reflow` 的 `reflow[]` 记录 `shape_id`、`action=move/resize/move_resize/unchanged`、`reason` 与 `from`/`to` 英寸坐标；无法安全放置时保留原状并返回 `warnings[type=unplaceable]`。
- `ppt_clean_unused_placeholders.py` 只删除空占位符，不删除空文本框。
- 美化后建议 `render_document_pages.py` 抽页检查（见 `render_qa.md`）。

## 框架图

`ppt_create_framework_diagram.py` 的 `FrameworkDiagram`：

```json
{
  "title": "系统架构",
  "stages": [
    {"label": "输入", "nodes": [{"id": "a", "text": "原始数据"}]},
    {"label": "处理", "nodes": [{"id": "b", "text": "特征提取"}, {"id": "c", "text": "模型推理"}]}
  ],
  "edges": [{"from_id": "a", "to_id": "b"}],
  "color_scheme": "professional-blue",
  "node_width": 1.4,
  "node_height": 0.55
}
```

- 左侧圆柱阶段标签 + 阶段背景容器；节点支持 `rect/rounded_rect/ellipse/diamond` 等。
- 未指定 `edges` 时相邻节点自动连箭头；`edge_style` 支持 `straight/elbow/curve`。
- 12 套配色复用 PPT 方案；`arrow_color` 留空用中性灰。
- `ppt_add_flowchart_slide.py` 是同一布局的 COM 活动文档版本，不启动 PowerPoint。
