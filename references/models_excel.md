# Excel JSON 模型

来源：`scripts/office_engine/models.py`（Pydantic v2）。
适用边界：Excel 相关 `--params` 模型；C# 后端按同一份 JSON 契约解析。

> 本文件由 models.py 生成；字段类型、默认值与说明以本表为准。

## `ColumnFormat`

Excel 列格式

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `number_format` | `str \| None` |  |  |
| `font_name` | `str \| None` |  |  |
| `font_size` | `int \| None` |  |  |
| `font_color` | `str \| None` |  |  |
| `bg_color` | `str \| None` |  |  |
| `bold` | `bool \| None` |  |  |
| `italic` | `bool \| None` |  |  |
| `wrap_text` | `bool \| None` |  |  |
| `width` | `float \| None` |  | 列宽 (字符数) |
| `alignment` | `Optional[Literal['left', 'center', 'right']]` |  |  |
| `vertical_alignment` | `Optional[Literal['top', 'center', 'bottom']]` |  |  |

## `ExcelChartDef`

Excel 图表定义（插入到指定工作表）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `chart_type` | `Literal['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'radar']` | 'column' | 图表类型 |
| `title` | `str \| None` |  |  |
| `data_range` | `str` | 必填 | 数据区域，如 'A1:C10' 或 'Sheet1!A1:C10'（首行/首列可作系列名与分类名） |
| `categories_range` | `str \| None` |  | 分类轴区域，如 'A2:A10'；缺省时由 data_range 推断 |
| `titles_from_data` | `bool` | true | data_range 首行是否作为系列名 |
| `anchor` | `str` | 'E2' | 图表左上角锚点单元格，如 'E2' |
| `width` | `float` | 15 | 图表宽度 (cm) |
| `height` | `float` | 8 | 图表高度 (cm) |
| `legend` | `Literal['right', 'bottom', 'top', 'left', 'none']` | 'right' | 图例位置 |
| `x_axis_title` | `str \| None` |  |  |
| `y_axis_title` | `str \| None` |  |  |
| `style` | `int \| None` |  | Excel 内置图表样式编号 1-48 |

## `ConditionalFormatDef`

Excel 条件格式规则

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `range` | `str` | 必填 | 作用区域，如 'A2:A100' |
| `type` | `Literal['cell_is', 'formula', 'color_scale', 'data_bar', 'icon_set']` | 'cell_is' | 规则类型 |
| `operator` | `Optional[Literal['greaterThan', 'lessThan', 'between', 'equal', 'notEqual', 'greaterThanOrEqual', 'lessThanOrEqual']]` |  | cell_is 比较运算符 |
| `formula` | `list[str] \| None` |  | 比较值或公式列表，如 ['100'] 或 ['$A1>5'] |
| `fill_color` | `str \| None` |  | 满足条件时的背景色 HEX（cell_is/formula） |
| `font_color` | `str \| None` |  | 满足条件时的字体色 HEX（cell_is/formula） |
| `min_color` | `str \| None` |  | 色阶最小值颜色（color_scale） |
| `mid_color` | `str \| None` |  | 色阶中间值颜色（color_scale） |
| `max_color` | `str \| None` |  | 色阶最大值颜色（color_scale） |
| `bar_color` | `str \| None` |  | 数据条颜色（data_bar） |
| `icon_style` | `str \| None` |  | 图标集名称（icon_set），如 '3TrafficLights1'，默认同左 |

## `DataValidationDef`

Excel 数据验证规则

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `range` | `str` | 必填 | 作用区域，如 'B2:B100' |
| `type` | `Literal['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom']` | 'list' | 验证类型 |
| `formula1` | `str` | 必填 | 条件 1；list 类型为逗号分隔值或区域引用，如 '"是,否"' 或 '=$D$1:$D$5' |
| `formula2` | `str \| None` |  | 条件 2（between 类验证使用） |
| `allow_blank` | `bool` | true | 是否允许空值 |
| `show_dropdown` | `bool` | true | list 类型是否显示下拉箭头 |
| `prompt_title` | `str \| None` |  |  |
| `prompt` | `str \| None` |  |  |
| `error_title` | `str \| None` |  |  |
| `error` | `str \| None` |  |  |

## `ExcelTableDef`

Excel 超级表（ListObject）定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `range` | `str` | 必填 | 表格区域，如 'A1:D10'（必须包含表头行） |
| `name` | `str` | 必填 | 表格名称（不含空格） |
| `style` | `str` | 'TableStyleMedium2' | 表格样式名 |
| `show_row_stripes` | `bool` | true |  |
| `show_totals_row` | `bool` | false |  |

## `DefinedNameDef`

Excel 定义名称

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str` | 必填 | 名称，如 'SalesTotal' |
| `refers_to` | `str` | 必填 | 引用内容，如 'Sheet1!$B$2:$B$10' 或 '=SUM(Sheet1!$B$2:$B$10)' |
| `local` | `bool` | false | True 表示只作用于当前工作表 |

## `ExcelHyperlinkDef`

Excel 单元格超链接

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `cell` | `str` | 必填 | 单元格，如 'A2' |
| `url` | `str` | 必填 | 链接地址 |
| `tooltip` | `str \| None` |  |  |

## `ExcelCommentDef`

Excel 单元格批注

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `cell` | `str` | 必填 | 单元格，如 'C3' |
| `text` | `str` | 必填 | 批注正文 |
| `author` | `str` | 'office-writer' | 批注作者 |

## `ExcelImageDef`

Excel 工作表图片

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `path` | `str` | 必填 | 本地图片路径 |
| `anchor` | `str` | 'A1' | 左上角锚点单元格 |
| `width` | `float \| None` |  | 宽度 (cm) |
| `height` | `float \| None` |  | 高度 (cm) |

## `ExcelPageSetupDef`

Excel 打印与页面设置

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `orientation` | `Optional[Literal['portrait', 'landscape']]` |  |  |
| `paper_size` | `Optional[Literal['A4', 'Letter', 'A3', 'B5']]` |  |  |
| `fit_to_width` | `int \| None` |  | 缩放到页宽 |
| `fit_to_height` | `int \| None` |  | 缩放到页高 |
| `print_area` | `str \| None` |  | 打印区域，如 'A1:F30' |
| `print_title_rows` | `str \| None` |  | 重复标题行，如 '1:1' |
| `print_title_cols` | `str \| None` |  | 重复标题列，如 'A:A' |
| `margin_top` | `float \| None` |  | 上边距 (英寸) |
| `margin_bottom` | `float \| None` |  | 下边距 (英寸) |
| `margin_left` | `float \| None` |  | 左边距 (英寸) |
| `margin_right` | `float \| None` |  | 右边距 (英寸) |
| `center_horizontally` | `bool \| None` |  |  |
| `center_vertically` | `bool \| None` |  |  |

## `SheetProtectionDef`

工作表保护（allow_* 为 True 表示允许该操作）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `password` | `str \| None` |  | 保护密码（明文，写入 OOXML hash） |
| `allow_select_locked_cells` | `bool` | true | 允许选中锁定单元格 |
| `allow_select_unlocked_cells` | `bool` | true | 允许选中未锁定单元格 |
| `allow_format_cells` | `bool` | false |  |
| `allow_format_columns` | `bool` | false |  |
| `allow_format_rows` | `bool` | false |  |
| `allow_insert_rows` | `bool` | false |  |
| `allow_insert_columns` | `bool` | false |  |
| `allow_delete_rows` | `bool` | false |  |
| `allow_delete_columns` | `bool` | false |  |
| `allow_sort` | `bool` | false |  |
| `allow_auto_filter` | `bool` | false |  |

## `SortKeyDef`

排序键：column 为表头名；has_header=false 时可用列字母，如 'B'

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `column` | `str` | 必填 | 表头名或列字母 |
| `ascending` | `bool` | true | 是否升序 |

## `ExcelSortDef`

按行排序（写数据后重排行，并写 sortState）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `range` | `str \| None` |  | 排序区域，如 'A1:D20'；默认整个已用区域 |
| `keys` | `list[SortKeyDef]` | 必填 | 排序键，按顺序比较 |
| `has_header` | `bool` | true | 首行是否为表头 |

## `SheetDef`

Excel 工作表定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str` | 必填 | 工作表名称 |
| `headers` | `list[str] \| None` |  | 表头行 |
| `data` | `list[list]` | 工厂 | 数据行（支持 str/int/float/None） |
| `allow_formulas` | `bool` | false | 是否将以 = 开头的字符串写为 Excel 公式；默认按纯文本写入 |
| `column_widths` | `list[float] \| None` |  | 列宽 (字符数) |
| `row_heights` | `dict[int, float] \| None` |  | 行高（行号→磅值），如 {"1": 24} |
| `header_style` | `CellStyle \| None` |  | 表头样式 |
| `column_formats` | `list[ColumnFormat] \| None` |  | 按列格式 |
| `merged_cells` | `list[MergeRange] \| None` |  | 合并单元格 |
| `freeze` | `str \| None` |  | 冻结窗格，如 'A2' 冻结首行 |
| `auto_filter` | `bool` | false | 是否启用自动筛选 |
| `tab_color` | `str \| None` |  | 工作表标签颜色 HEX |
| `sheet_state` | `Literal['visible', 'hidden', 'veryHidden']` | 'visible' | 工作表可见性 |
| `charts` | `list[ExcelChartDef] \| None` |  | 插入的图表 |
| `conditional_formats` | `list[ConditionalFormatDef] \| None` |  | 条件格式规则 |
| `data_validations` | `list[DataValidationDef] \| None` |  | 数据验证规则 |
| `tables` | `list[ExcelTableDef] \| None` |  | 超级表 |
| `defined_names` | `list[DefinedNameDef] \| None` |  | 定义名称（工作簿级或本表级） |
| `hyperlinks` | `list[ExcelHyperlinkDef] \| None` |  | 单元格超链接 |
| `comments` | `list[ExcelCommentDef] \| None` |  | 单元格批注 |
| `images` | `list[ExcelImageDef] \| None` |  | 插入图片 |
| `page_setup` | `ExcelPageSetupDef \| None` |  | 打印与页面设置 |
| `protection` | `SheetProtectionDef \| None` |  | 工作表保护 |
| `sort` | `ExcelSortDef \| None` |  | 写数据后按行排序 |

## `ExcelContent`

Excel 工作簿内容

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `sheets` | `list[SheetDef]` | 必填 | 工作表列表（至少 1 个） |
