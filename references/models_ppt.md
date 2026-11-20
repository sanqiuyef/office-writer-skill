# PowerPoint JSON 模型

来源：`scripts/office_engine/models.py`（Pydantic v2）。
适用边界：PowerPoint 相关 `--params` 模型；C# 后端按同一份 JSON 契约解析。

> 本文件由 models.py 生成；字段类型、默认值与说明以本表为准。

## `TextStyle`

PPT 文字样式

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `font_name` | `str \| None` |  |  |
| `font_size` | `int \| None` |  |  |
| `bold` | `bool \| None` |  |  |
| `italic` | `bool \| None` |  |  |
| `color` | `str \| None` |  |  |
| `alignment` | `Optional[Literal['left', 'center', 'right']]` |  |  |

## `SlideTableDef`

PPT 幻灯片表格定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `headers` | `list[str]` | 工厂 |  |
| `rows` | `list[list[str]]` | 工厂 |  |
| `x` | `float \| None` |  | 左侧位置 (英寸)，默认 0.5 |
| `y` | `float \| None` |  | 顶部位置 (英寸)，默认 2.0 |
| `width` | `float \| None` |  | 表格宽度 (英寸)，默认 7.0 |
| `height` | `float \| None` |  | 表格高度 (英寸)，默认 2.5 |
| `column_widths` | `list[float] \| None` |  | 各列宽度 (英寸)，总宽超出表格宽度时按比例缩放 |
| `header_bold` | `bool` | True | 表头是否加粗 |

## `SlideImageDef`

PPT 幻灯片图片定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `path` | `str \| None` |  | 图片文件路径（与 url 二选一） |
| `url` | `str \| None` |  | 图片 URL（与 path 二选一），自动下载后插入 |
| `x` | `float \| None` |  | 左侧位置 (英寸) |
| `y` | `float \| None` |  | 顶部位置 (英寸) |
| `width` | `float \| None` |  | 宽度 (英寸) |
| `height` | `float \| None` |  | 高度 (英寸) |
| `crop_left` | `float \| None` |  | 左侧裁剪比例 (0~1) |
| `crop_right` | `float \| None` |  | 右侧裁剪比例 (0~1) |
| `crop_top` | `float \| None` |  | 顶部裁剪比例 (0~1) |
| `crop_bottom` | `float \| None` |  | 底部裁剪比例 (0~1) |
| `hyperlink_url` | `str \| None` |  | 点击图片打开的链接 |

## `SlideShapeDef`

PPT 形状定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `shape_type` | `Literal['rect', 'rounded_rect', 'oval', 'arrow', 'cylinder', 'diamond']` | 'rect' |  |
| `text` | `str \| None` |  | 形状内文字 |
| `hyperlink_url` | `str \| None` |  | 点击形状打开的链接 |
| `x` | `float \| None` |  | 左侧位置 (英寸) |
| `y` | `float \| None` |  | 顶部位置 (英寸) |
| `width` | `float` | 必填 | 宽度 (英寸) |
| `height` | `float` | 必填 | 高度 (英寸) |
| `bg_color` | `str \| None` |  | 填充颜色 HEX，None 且 no_fill=False 时保持默认 |
| `font_color` | `str \| None` |  | 文字颜色 HEX |
| `font_size` | `int \| None` |  | 字号 (pt)，默认 12 |
| `bold` | `bool \| None` | False |  |
| `no_fill` | `bool` | False | 无填充（透明），用于虚线分组框等 |
| `line_color` | `str \| None` |  | 边框颜色 HEX，None 不设置边框 |
| `line_width` | `float \| None` |  | 边框宽度 (pt) |
| `line_dash` | `bool` | False | 边框是否虚线 |
| `vertical_text` | `bool` | False | 文字竖排（圆柱体阶段标记用） |

## `SlideConnectorDef`

PPT 连接线/箭头定义（坐标单位：英寸，相对幻灯片左上角）

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `x1` | `float` | 必填 | 起点 X (英寸) |
| `y1` | `float` | 必填 | 起点 Y (英寸) |
| `x2` | `float` | 必填 | 终点 X (英寸) |
| `y2` | `float` | 必填 | 终点 Y (英寸) |
| `color` | `str \| None` |  | 线条颜色 HEX，留空用默认 |
| `width` | `float` | 1.5 | 线宽 (pt) |
| `dash` | `bool` | False | 是否虚线（损失反向传播等特殊箭头） |
| `arrowhead` | `bool` | True | 终点是否带箭头 |
| `arrowhead_style` | `Literal['triangle', 'open', 'stealth']` | 'triangle' | 箭头样式 |
| `begin_arrowhead` | `bool` | False | 起点是否带箭头 |

## `SlideItem`

幻灯片中的一个内容元素

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `type` | `Literal['paragraph', 'table', 'image', 'shape', 'connector', 'bullet_list', 'numbered_list', 'chart']` | 必填 |  |
| `text` | `str \| None` |  |  |
| `level` | `int` | 0 | 缩进层级 (0=顶级) |
| `style` | `TextStyle \| None` |  |  |
| `table` | `SlideTableDef \| None` |  |  |
| `image` | `SlideImageDef \| None` |  |  |
| `shape` | `SlideShapeDef \| None` |  |  |
| `connector` | `SlideConnectorDef \| None` |  | 连接线定义，仅 connector 类型使用 |
| `items` | `list[str] \| None` |  | 列表项，仅 bullet_list/numbered_list 使用 |
| `chart` | `ForwardRef('ChartDef \| None')` |  | 图表定义，仅 chart 类型使用 |

## `ChartDef`

PPT 图表定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `chart_type` | `Literal['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'radar']` | 'column' | 图表类型 |
| `categories` | `list[str]` | 必填 | 类别标签（X 轴） |
| `series` | `ForwardRef('list[ChartSeries]')` | 必填 | 数据系列 |
| `has_legend` | `bool` | True | 是否显示图例 |
| `title` | `str \| None` |  | 图表标题 |
| `x` | `float \| None` |  | 左侧位置 (英寸)，默认 0.5 |
| `y` | `float \| None` |  | 顶部位置 (英寸)，默认 1.5 |
| `width` | `float \| None` |  | 宽度 (英寸)，默认 8.0 |
| `height` | `float \| None` |  | 高度 (英寸)，默认 4.5 |

## `ChartSeries`

图表数据系列

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str` | 必填 | 系列名称 |
| `values` | `list[float]` | 必填 | 数值列表 |

## `SlideDef`

幻灯片定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `layout` | `Literal['title', 'section', 'content', 'two_content', 'comparison', 'title_only', 'blank', 'content_caption', 'picture_caption']` | 'content' |  |
| `title` | `str \| None` |  | 幻灯片标题 |
| `subtitle` | `str \| None` |  | 副标题（仅 title/section 布局） |
| `body` | `list[SlideItem] \| None` |  | 主体内容 |
| `left` | `list[SlideItem] \| None` |  | 左侧内容（仅 two_content/comparison 布局） |
| `right` | `list[SlideItem] \| None` |  | 右侧内容（仅 two_content/comparison 布局） |
| `notes` | `str \| None` |  | 演讲者备注 |
| `bg_color` | `str \| None` |  | 幻灯片背景色 HEX |
| `slide_width` | `float \| None` |  | 幻灯片宽度 (英寸)，默认 10 |
| `slide_height` | `float \| None` |  | 幻灯片高度 (英寸)，默认 7.5 |

## `PptContent`

PPT 演示文稿内容

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `slides` | `list[SlideDef]` | 必填 | 幻灯片列表（至少 1 个） |

## `FwNode`

框架图节点

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `id` | `str` | 必填 | 节点唯一标识 |
| `text` | `str` | 必填 | 节点显示文本，支持 \n 换行 |
| `style` | `Literal['rect', 'rounded_rect', 'ellipse', 'diamond', 'process', 'decision', 'data', 'document', 'terminator', 'predefined', 'cylinder', 'parallelogram', 'chevron', 'right_arrow', 'left_arrow', 'callout', 'cloud']` | 'rounded_rect' | 节点形状 |
| `fill_color` | `str \| None` |  | 填充色 HEX，留空使用阶段色 |
| `font_size` | `int` | 12 | 字号 |
| `bold` | `bool` | True | 是否加粗 |

## `FwEdge`

框架图连接线

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `from_id` | `str` | 必填 | 起始节点 ID |
| `to_id` | `str` | 必填 | 目标节点 ID |
| `label` | `str \| None` |  | 箭头上的标签文字（可选） |
| `edge_style` | `Literal['straight', 'elbow', 'curve']` | 'straight' | 连接线样式：straight(直线)/elbow(正交)/curve(曲线) |

## `FwStage`

框架图阶段分组

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `label` | `str` | 必填 | 阶段标签，显示在左侧圆柱形上 |
| `color` | `str \| None` |  | 阶段背景色 HEX，留空使用配色方案默认 |
| `nodes` | `list[FwNode]` | 工厂 | 阶段内的节点列表 |

## `FrameworkDiagram`

框架图完整定义

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `title` | `str` | '' | 框架图总标题 |
| `subtitle` | `str \| None` |  | 框架图副标题 |
| `stages` | `list[FwStage]` | 必填 | 阶段列表，按从上到下顺序排列 |
| `edges` | `list[FwEdge]` | 工厂 | 节点间连接线，未指定的相邻节点自动生成箭头 |
| `color_scheme` | `str` | 'professional-blue' | 配色方案名称 |
| `node_width` | `float` | 1.4 | 默认节点宽度 (英寸) |
| `node_height` | `float` | 0.55 | 默认节点高度 (英寸) |
| `stage_gap` | `float` | 0.15 | 阶段间垂直间距 (英寸) |
| `arrow_color` | `str \| None` |  | 箭头颜色 HEX |
| `font_cn` | `str` | '宋体' | 中文字体 |
| `font_en` | `str` | 'Times New Roman' | 西文字体 |
| `font_size` | `int` | 16 | 正文默认字号 (pt)，放不下时自动缩小 |
| `title_font_size` | `int` | 24 | 标题字号 (pt) |

## `ColorScheme`

PPT 配色方案

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `name` | `str` | 必填 | 方案名称标识 |
| `display_name` | `str` | '' | 方案显示名称 |
| `description` | `str \| None` |  | 方案描述 |
| `title_color` | `str` | 必填 | 标题颜色 HEX |
| `body_color` | `str` | 必填 | 正文颜色 HEX |
| `accent_color` | `str` | 必填 | 强调色 HEX |
| `background_color` | `str` | 必填 | 幻灯片背景色 HEX |
| `shape_fill_color` | `str` | 必填 | 形状填充色 HEX |
| `table_header_color` | `str` | 必填 | 表格表头背景色 HEX |
| `title_font` | `str` | '微软雅黑' | 标题字体 |
| `body_font` | `str` | '微软雅黑' | 正文字体 |
| `title_size` | `int` | 32 | 标题字号 (pt) |
| `body_size` | `int` | 18 | 正文字号 (pt) |

## `TemplateAnalysis`

PPT 模板分析结果

| 字段 | 类型 | 默认 | 说明 |
|:---|:---|:---|:---|
| `file_path` | `str` | 必填 | 模板文件路径 |
| `slide_count` | `int` | 0 | 幻灯片数量 |
| `slide_width` | `float` | 0 | 幻灯片宽度 (英寸) |
| `slide_height` | `float` | 0 | 幻灯片高度 (英寸) |
| `layout_names` | `list[str]` | 工厂 | 可用布局名称列表 |
| `color_schemes` | `list[ColorScheme] \| None` |  | 提取的配色方案 |
| `fonts_used` | `list[str]` | 工厂 | 模板中使用的字体 |
