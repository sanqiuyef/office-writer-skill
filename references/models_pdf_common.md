# PDF 与通用模型

来源：`scripts/office_engine/models.py`（Pydantic v2）。
适用边界：PDF/通用 相关 `--params` 模型；C# 后端按同一份 JSON 契约解析。

> 本文件由 models.py 生成；字段类型、默认值与说明以本表为准。

## `PdfTableDef`

PDF 表格定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `headers` | `list[str]` | 工厂 | 表头行 |
| `rows` | `list[list[str]]` | 工厂 | 数据行 |
| `column_widths` | `list[float] \| None` |  | 列宽比例 (0~1)，如 [0.3, 0.4, 0.3] |

## `PdfSectionItem`

PDF 文档的一个内容块

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `type` | `Literal['heading', 'paragraph', 'table', 'image', 'page_break', 'bullet_list', 'numbered_list']` | 必填 |  |
| `text` | `str \| None` |  | 文字内容 |
| `level` | `int \| None` |  | 标题层级 (1-6)，仅 heading 类型使用 |
| `style` | `ParagraphStyle \| None` |  | 段落样式 |
| `items` | `list[str] \| None` |  | 列表项，仅 bullet_list/numbered_list 使用 |
| `table` | `PdfTableDef \| None` |  | 表格定义，仅 table 类型使用 |
| `image` | `ImageDef \| None` |  | 图片定义，仅 image 类型使用 |

## `PdfContent`

PDF 文档内容

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `sections` | `list[PdfSectionItem]` | 必填 | 文档内容块列表 |

## `PdfPageSetup`

PDF 页面设置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `paper_size` | `Literal['A4', 'Letter', 'A3', 'B5']` | 'A4' |  |
| `orientation` | `Literal['portrait', 'landscape']` | 'portrait' |  |
| `margin_top` | `float` | 10 | 上边距 (mm) |
| `margin_bottom` | `float` | 10 | 下边距 (mm) |
| `margin_left` | `float` | 10 | 左边距 (mm) |
| `margin_right` | `float` | 10 | 右边距 (mm) |

## `PdfFontConfig`

PDF 字体配置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `cn_font` | `str` | '' | 中文字体名称，如 'simhei'、'simsun'，留空自动检测 |
| `cn_font_path` | `str \| None` |  | 自定义中文字体文件路径 (.ttf/.ttc) |
| `en_font` | `str` | 'Helvetica' | 西文字体: Helvetica/Courier/Times |
| `cn_font_size` | `int \| str` | '小四' | 正文字号 |
| `heading_font_size` | `int \| str` | '三号' | 标题字号 |

## `CreateResult`

创建操作的返回结果

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `success` | `bool` | 必填 |  |
| `file_path` | `str` | 必填 |  |
| `file_size` | `int` | 0 |  |
| `sheets_count` | `int \| None` |  |  |
| `slides_count` | `int \| None` |  |  |
| `message` | `str` | 必填 |  |
| `backup_path` | `str \| None` |  |  |

## `EditResult`

编辑操作的返回结果

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `success` | `bool` | 必填 |  |
| `file_path` | `str` | 必填 |  |
| `operation_count` | `int` | 0 |  |
| `message` | `str` | 必填 |  |
| `backup_path` | `str \| None` |  |  |

## `OfficeOperation`

A path-addressed Office document mutation used by batch editing.

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `op` | `Literal['set', 'add', 'remove', 'move', 'replace']` | 必填 |  |
| `path` | `str` | 必填 | 目标路径，如 /body/p[1]、/Sheet1/A1 或 /slide[1]/shape[2] |
| `element_type` | `str \| None` |  | add 操作的元素类型，如 paragraph、sheet、slide、textbox |
| `props` | `dict[str, Any]` | 工厂 | 操作属性 |
| `target_path` | `str \| None` |  | move 操作的目标路径 |
| `position` | `Literal['before', 'after']` | 'after' | move 相对目标的位置 |

## `ReplaceSection`

Word 活动文档的按标题替换节。

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `heading_text` | `str` | 必填 | 定位标题文字 |
| `heading_match` | `Literal['exact', 'contains']` | 'contains' | 标题匹配方式 |
| `end_heading_text` | `str \| None` |  | 结束标题文字（不含）；为空则到下一个同级或更高级标题 |
| `keep_heading` | `bool` | True | 是否保留标题段落 |
| `lines` | `list[str]` | 工厂 | 替换后的内容行，每行一段 |
| `font` | `FontConfig \| None` |  | 新内容字体 |

## `TextReplacement`

活动文档查找替换（Word/PPT/Excel 通用）。

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `find` | `str` | 必填 | 待查找文本 |
| `replace` | `str` | '' | 替换文本 |
| `match_case` | `bool` | False | 是否区分大小写 |

## `WordCommentDef`

Word 批注（锚定到正文文字或段落）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `text` | `str` | 必填 | 批注正文 |
| `author` | `str` | 'office-writer' | 批注作者 |
| `initials` | `str` | 'OW' | 作者缩写 |
| `anchor_text` | `str \| None` |  | 锚定文字：命中的第一处 run 上添加批注 |
| `paragraph_index` | `int \| None` |  | 锚定到第 N 个正文段落（1 起）；与 anchor_text 二选一 |
| `match_case` | `bool` | False | anchor_text 是否区分大小写 |

## `PptHyperlinkDef`

PPT 超链接（按形状或形状内文字定位）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `slide` | `int` | 必填 | 幻灯片序号（1 起） |
| `url` | `str` | 必填 | 链接地址 |
| `shape_index` | `int \| None` |  | 形状序号（1 起，对应 inspect 路径 /slide[n]/shape[i]） |
| `shape_text` | `str \| None` |  | 按形状文字包含匹配定位 |
| `text_match` | `str \| None` |  | 仅给形状内匹配的文字 run 加链接；留空则整个形状可点击 |
| `match_case` | `bool` | False | text_match 是否区分大小写 |
