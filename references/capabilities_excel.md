# Excel 能力矩阵

来源：Microsoft Excel VBA 对象模型（learn.microsoft.com/office/vba/api/excel.*）、openpyxl 3.1.5、本 skill 的 C# office-native 引擎实现（2026-11）。
适用边界：文件级 .xlsx/.xlsm 读写与 Windows Excel/WPS COM 转换；不含 Excel 网页版、Power Query 编辑器、VBA 宏执行。

状态记号：**已覆盖** = 有入口且已实测；**部分** = 能力有限制；**未覆盖** = 当前没有入口，括号内给替代路径。

| 功能域 | 能力项 | 状态 | 入口 / 参数 / 限制 |
|:---|:---|:---|:---|
| 工作簿/工作表 | 新建多表、模板新建、追加工作表、增删/排序、隐藏/veryHidden、标签色 | 已覆盖 | `create_excel_workbook`、`edit_excel_workbook`、`batch_edit_office_document`；`SheetDef.sheet_state` |
| 模板 | 模板新建后清理失效本地定义名、重算 activeTab；`{{占位符}}` 合并值以 `=` 开头时写成公式 | 已覆盖 | `--template`、`merge_office_template`；合并后的公式不预计算缓存值 |
| 单元格/区域 | 值、公式（显式开关）、合并、增删行列、移动；批量写入合并区非锚点会明确报错且不落盘 | 已覆盖 | `SheetDef.data` + `allow_formulas`、`merged_cells`；批量 `set/add/remove/move` |
| 读取 | 值、公式缓存、日期/时间/时长格式、布尔、数字格式；`max_rows_per_sheet` 流式截断 | 已覆盖 | `read_excel_workbook`；日期按 openpyxl `data_only` 语义输出，无缓存公式返回空串 |
| 样式 | 字体名/字号/粗体/斜体/颜色、纯色填充、水平/垂直对齐、换行、数字格式、四周边框 | 已覆盖 | `CellStyle`、`ColumnFormat` |
| 布局 | 列宽、行高、冻结窗格、自动筛选 | 已覆盖 | `column_widths`、`ColumnFormat.width`、`row_heights`、`freeze`、`auto_filter` |
| 图表 | column/bar/line/pie/doughnut/area/scatter/radar，标题、图例、轴标题、样式、锚点、宽高 | 已覆盖 | `SheetDef.charts`（创建时）或 `excel_add_chart`（已有文件）；scatter 无 `categories_range` 时 data_range 必须恰好两列（X,Y），有 `categories_range` 时 data_range 的每一列都是一个 Y 系列 |
| 条件格式 | cell_is、formula、color_scale、data_bar、icon_set | 已覆盖 | `SheetDef.conditional_formats`；图标集用内置样式名（默认 3TrafficLights1） |
| 数据验证 | list / whole / decimal / date / time / textLength / custom，提示与错误信息 | 已覆盖 | `SheetDef.data_validations`；`show_dropdown` 控制下拉箭头 |
| 超级表 | 表名、样式、行条纹、总计行；表名与定义名冲突、表间范围相交、覆盖合并区会明确报错 | 已覆盖 | `SheetDef.tables`；总计行会把 ref 扩一行并设 `totalsRowCount=1` |
| 定义名称 | 工作簿级、工作表级；模板新建会清理失效本地名并重算 activeTab | 已覆盖 | `SheetDef.defined_names`；同名（忽略大小写）明确报错，不静默覆盖；局部名称写入 `workbook.xml`，openpyxl 回读不显示（Excel 可见） |
| 批注 | 传统单元格批注 | 已覆盖 | `SheetDef.comments`；无线程批注 |
| 超链接 | 单元格外部链接、tooltip | 已覆盖 | `SheetDef.hyperlinks`；只改字体颜色/下划线，保留原样式 |
| 图片 | 按锚点插入、cm 尺寸 | 已覆盖 | `SheetDef.images`；cm 按 96 DPI 换算像素 |
| 页面设置 | 方向、纸张、缩放、打印区域、重复标题行/列、页边距、居中 | 已覆盖 | `SheetDef.page_setup` |
| 公式计算 | 写入公式、用 dsoffice/Excel/WPS COM 重算并写缓存值 | 已覆盖 | `recalculate_excel_formulas`；openpyxl 本身不计算 |
| 导出/渲染 | xlsx → PDF、逐页/指定区域 PNG | 已覆盖 | `convert_to_pdf`、`render_document_pages` |
| 数据透视表 | 创建、改字段、刷新 | 已覆盖 | `excel_manage_pivots`（office-native COM）：create/list/refresh；字段支持行/列/值聚合 |
| 真实排序 | 按值排序并改变行顺序 | 未覆盖 | `auto_filter` 只是筛选；数据排序用 pandas，版式写入仍用本 skill |
| 形状/文本框/绘图 | 插入形状、绘图对象 | 未覆盖 | openpyxl 读取时会丢失 shapes/drawings |
| 迷你图/切片器/日程表 | sparkline、slicer、timeline | 未覆盖 | openpyxl 读取时移除扩展；xlsxwriter 可写迷你图但只能新建 |
| 宏/VBA | 保留、编辑、运行 | 部分 | `.xlsm/.xltm` 通过 `keep_vba` 保留；不能编辑或运行宏 |
| 密码/加密/签名 | 打开密码、加密、数字签名 | 未覆盖 | openpyxl 无解密/签名实现 |
| 旧格式/其他格式 | .xls / .xlsb / .ods / CSV | 部分 | `.xls` 仅转换/COM；`.xlsb/.ods` 未覆盖，CSV 用 pandas |
| Power Query/数据模型 | 查询、Power Pivot、关系 | 未覆盖 | 需要 COM（Workbook.Queries/Model）或 C# |
| 保护 | 工作表/工作簿保护 | 未覆盖 | openpyxl 可写保护，但本 skill 未封装入口 |
| 打印 | 导出 PDF、PrintOut | 已覆盖 | `excel_print`（office-native COM）：pdf/print；print 必须显式 `confirm_print=true` |

## 活动文档边界

`edit_open_excel_workbook` 只支持旧字段（表头/数据/列宽/合并/冻结/筛选）。图表、条件格式、数据验证、超级表、定义名称、批注、超链接、图片、页面设置、行高、隐藏表、边框/斜体等扩展字段会**明确报错**（退出码 3），不会静默丢弃；请关闭文件后改用文件级 `edit_excel_workbook`。

## 使用建议

标注 office-native 的能力默认走 C#（Windows + Excel）；未构建时入口返回退出码 3 并给出构建命令，或设 `OFFICE_NATIVE_ENGINE=python` 回退旧引擎。

- 新建报表：`create_excel_workbook` 一次写数据、样式、图表、条件格式、验证、超级表与页面设置。
- 已有文件加图表：`excel_add_chart`（原地写回并自动 `.bak`）。
- 需要透视表、打印、真实排序、宏：透视/打印用 `excel_manage_pivots`/`excel_print`（默认 C# COM）；排序用 pandas/openpyxl，宏需人工在 Excel 中处理。
