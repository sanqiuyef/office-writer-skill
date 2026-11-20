# Word JSON 模型

来源：`scripts/office_engine/models.py`（Pydantic v2）。
适用边界：Word 相关 `--params` 模型；C# 后端按同一份 JSON 契约解析。

> 本文件由 models.py 生成；字段类型、默认值与说明以本表为准。

## `FontConfig`

字体配置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str \| None` |  | 字体名称，如 '宋体'、'微软雅黑'、'Arial' |
| `size` | `int \| str \| None` |  | 字号（国内标准如'小四''五号'，或直接 pt 数值） |
| `bold` | `bool \| None` |  | 是否加粗 |
| `italic` | `bool \| None` |  | 是否斜体 |
| `color` | `str \| None` |  | 字体颜色 (HEX, 如 'FF0000') |

## `BorderSpec`

单元格四边框样式（四边同款）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `style` | `Literal['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair']` | 'thin' |  |
| `color` | `str \| None` |  | 边框颜色 HEX，默认黑色 |

## `CellStyle`

单元格样式

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `bold` | `bool \| None` |  |  |
| `italic` | `bool \| None` |  |  |
| `font_name` | `str \| None` |  |  |
| `font_size` | `int \| str \| None` |  |  |
| `font_color` | `str \| None` |  | 字体颜色 HEX |
| `bg_color` | `str \| None` |  | 背景颜色 HEX |
| `number_format` | `str \| None` |  | 数字格式，如 '#,##0.00'、'0%'、'YYYY-MM-DD' |
| `alignment` | `Optional[Literal['left', 'center', 'right']]` |  |  |
| `vertical_alignment` | `Optional[Literal['top', 'center', 'bottom']]` |  |  |
| `wrap_text` | `bool \| None` |  |  |
| `border` | `BorderSpec \| None` |  | 四周边框；None 不设置 |

## `PageSetup`

页面设置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `orientation` | `Optional[Literal['portrait', 'landscape']]` |  |  |
| `paper_size` | `Optional[Literal['A4', 'Letter', 'A3', 'B5']]` | 'A4' |  |
| `margin_top` | `float \| None` |  | 上边距 (英寸) |
| `margin_bottom` | `float \| None` |  | 下边距 (英寸) |
| `margin_left` | `float \| None` |  | 左边距 (英寸) |
| `margin_right` | `float \| None` |  | 右边距 (英寸) |
| `columns` | `int \| None` |  | 分栏数（1-4） |
| `column_spacing` | `float \| None` |  | 分栏间距 (英寸) |
| `page_border` | `Optional[Literal['single', 'double', 'dashed', 'dotted']]` |  | 页面边框线型 |
| `page_border_color` | `str \| None` |  | 页面边框颜色 HEX |
| `watermark_text` | `str \| None` |  | 文字水印内容 |
| `watermark_color` | `str \| None` |  | 水印颜色 HEX，默认浅灰 |

## `HeaderFooter`

页眉页脚设置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `header_text` | `str \| None` |  | 页眉文字 |
| `footer_text` | `str \| None` |  | 页脚文字，用 {page} 代表页码 |

## `SceneProfile`

场景排版配置文件

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str` | 必填 | 场景名称标识 |
| `display_name` | `str` | '' | 场景显示名称 |
| `description` | `str \| None` |  | 场景描述 |
| `font_name` | `str` | '宋体' |  |
| `font_name_western` | `str` | 'Times New Roman' | 西文字体 |
| `heading_font` | `str \| None` |  |  |
| `body_size` | `str` | '小四' |  |
| `heading_sizes` | `dict[int, str]` | 工厂 |  |
| `line_spacing` | `float` | 1.3 |  |
| `first_line_indent` | `int \| None` | 28 |  |
| `alignment` | `str` | 'justify' |  |
| `paper_size` | `str` | 'A4' |  |
| `margin_top` | `float` | 1.0 |  |
| `margin_bottom` | `float` | 1.0 |  |
| `margin_left` | `float` | 1.18 |  |
| `margin_right` | `float` | 1.18 |  |
| `page_number_style` | `Literal['none', 'roman', 'arabic']` | 'arabic' |  |
| `show_cover` | `bool` | false |  |

## `CoverConfig`

封面配置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `title` | `str \| None` |  | 封面主标题；留空时用第一个 heading 或默认标题 |
| `subtitle` | `str \| None` |  | 封面副标题 |
| `author` | `str \| None` |  | 作者/创建者 |
| `date` | `str \| None` |  | 日期 |
| `organization` | `str \| None` |  | 机构/单位名称 |
| `extra_lines` | `list[str] \| None` |  | 附加信息行 |
| `recipe_id` | `str \| None` |  | 封面配方 ID (R1-R7)，自动选择时可留空 |
| `doc_type` | `str \| None` |  | 文档类型（general/report/thesis 等），用于自动选配方 |

## `MergeRange`

表格合并单元格范围 (Excel A1 风格)

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `range` | `str` | 必填 | 合并范围，如 'A1:B2'、'C1:C3' |

## `ParagraphStyle`

段落样式

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `font_name` | `str \| None` |  |  |
| `font_size` | `int \| str \| None` |  | 字号（国内标准如'小四''五号'，或直接 pt 数值） |
| `bold` | `bool \| None` |  |  |
| `italic` | `bool \| None` |  |  |
| `color` | `str \| None` |  |  |
| `alignment` | `Optional[Literal['left', 'center', 'right', 'justify']]` |  |  |
| `line_spacing` | `float \| None` |  | 行距倍数，如 1.0、1.5、2.0 |
| `first_line_indent` | `int \| None` |  | 首行缩进 (pt)，中文段落常用 28pt≈2字符 |
| `space_before` | `int \| None` |  | 段前距 (pt) |
| `space_after` | `int \| None` |  | 段后距 (pt) |

## `TableDef`

Word 表格定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `headers` | `list[str]` | 工厂 | 表头行 |
| `rows` | `list[list[str]]` | 工厂 | 数据行 |
| `column_widths` | `list[float] \| None` |  | 列宽 (英寸) |
| `header_style` | `CellStyle \| None` |  | 表头样式 |
| `merged_cells` | `list[MergeRange] \| None` |  | 合并单元格 |

## `ImageDef`

图片定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `path` | `str` | 必填 | 图片文件路径 |
| `width` | `float \| None` |  | 宽度 (英寸) |
| `height` | `float \| None` |  | 高度 (英寸) |

## `HyperlinkDef`

Word 超链接定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `text` | `str` | 必填 | 显示文字 |
| `url` | `str` | 必填 | 链接地址（http/https/mailto/file） |
| `tooltip` | `str \| None` |  | 悬停提示 |

## `FieldDef`

Word 域定义（PAGE/NUMPAGES/DATE 等）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `field_type` | `Literal['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME', 'FILENAME', 'AUTHOR', 'TITLE', 'CUSTOM']` | 'PAGE' | 域类型；CUSTOM 时用 code 提供完整域代码 |
| `code` | `str \| None` |  | 自定义域代码，仅 CUSTOM 类型使用，如 ' MERGEFIELD 姓名 ' |
| `fallback_text` | `str \| None` |  | 域结果占位文字，Word 更新域前显示 |
| `number_format` | `str \| None` |  | 数字格式开关，如 'roman'、'DBNUM1'；用于 PAGE/NUMPAGES |

## `BookmarkDef`

Word 书签（包裹 text 内容）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str` | 必填 | 书签名，字母/数字/下划线，不超过 40 字符 |

## `SectionItem`

Word 文档的一个内容块

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `type` | `Literal['heading', 'paragraph', 'table', 'image', 'page_break', 'bullet_list', 'numbered_list', 'code', 'toc', 'section_break', 'hyperlink', 'field', 'bookmark']` | 必填 |  |
| `text` | `str \| None` |  | 文字内容 |
| `level` | `int \| None` |  | 标题层级 (1-9)，仅 heading 类型使用 |
| `language` | `str \| None` |  | 代码语言标签，仅 code 类型使用 |
| `style` | `ParagraphStyle \| None` |  | 段落样式 |
| `items` | `list[str] \| None` |  | 列表项，仅 bullet_list/numbered_list 使用 |
| `table` | `TableDef \| None` |  | 表格定义，仅 table 类型使用 |
| `image` | `ImageDef \| None` |  | 图片定义，仅 image 类型使用 |
| `toc_levels` | `str \| None` |  | TOC 包含的标题层级，如 '1-3'，仅 toc 类型使用 |
| `hyperlink` | `HyperlinkDef \| None` |  | 超链接定义，仅 hyperlink 类型使用 |
| `field` | `FieldDef \| None` |  | 域定义，仅 field 类型使用 |
| `bookmark` | `BookmarkDef \| None` |  | 书签定义，仅 bookmark 类型使用 |

## `WordProperties`

Word 文档核心属性

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `title` | `str \| None` |  |  |
| `subject` | `str \| None` |  |  |
| `author` | `str \| None` |  |  |
| `keywords` | `str \| None` |  |  |
| `comments` | `str \| None` |  |  |
| `category` | `str \| None` |  |  |
| `last_modified_by` | `str \| None` |  |  |
| `created` | `str \| None` |  | 创建时间，ISO 8601，如 '2026-11-19T10:00:00' |
| `modified` | `str \| None` |  | 修改时间，ISO 8601 |

## `WordContent`

Word 文档内容

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `sections` | `list[SectionItem]` | 必填 | 文档内容块列表 |
| `properties` | `WordProperties \| None` |  | 写入文档核心属性（标题/作者/主题等） |
