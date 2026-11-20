# PowerPoint 能力矩阵

来源：Microsoft PowerPoint VBA 对象模型（learn.microsoft.com/office/vba/api/powerpoint.*）、python-pptx 1.0.2、本 skill 的 C# office-native 引擎实现（2026-11）。
适用边界：文件级 .pptx/.potx 编辑与 Windows PowerPoint/WPS COM；不含 .ppt 文件级编辑；动画/切换/批注/SmartArt 走 COM（office-native）。

状态记号：**已覆盖** = 有入口且已实测；**部分** = 能力有限制；**未覆盖** = 当前没有入口，括号内给替代路径。

| 功能域 | 能力项 | 状态 | 入口 / 参数 / 限制 |
|:---|:---|:---|:---|
| 演示文稿 | 新建、模板（含 .potx）、追加幻灯片、删除/移动幻灯片 | 已覆盖 | `create_powerpoint`、`edit_powerpoint`、`batch_edit_office_document`；.potx 用临时工作副本改内容类型 |
| 版式 | 9 种内置版式（title/section/content/two_content/comparison/title_only/blank/content_caption/picture_caption） | 已覆盖 | `SlideDef.layout`；按版式名映射固定索引，自定义模板布局顺序不同会错位 |
| 占位符 | 标题/副标题/正文填充、清空空白占位符、双栏左右分栏 | 已覆盖 | `_build_slide_body(placeholder_idx=1/2)`；无新增任意占位符 API |
| 文本 | 段落、级别、字体名/字号/粗斜体/颜色、对齐、中英混排 | 已覆盖 | `SlideItem.type=paragraph` + `TextStyle` |
| 列表 | 项目符号、编号列表（真实 `a:buChar` / `a:buAutoNum`） | 已覆盖 | `bullet_list`、`numbered_list`；无多级自定义编号 |
| 表格 | 表头/数据行、位置、宽高、列宽、表头加粗 | 已覆盖 | `SlideTableDef`；无单元格边框/底纹 API |
| 图表 | column/bar/line/pie/doughnut/area/radar，标题、图例、位置尺寸；内嵌可编辑数据工作簿（`c:externalData` + embedded xlsx） | 已覆盖 | `SlideItem.type=chart` + `ChartDef`；无 trendline/error bar/组合图 |
| 批量编辑 | 按 stable path 设置/增删/移动/替换；跨 slide 移动图片/图表/超链接会迁移关系 | 已覆盖 | `batch_edit_office_document`；音视频等媒体关系跨 slide 明确不支持 |
| 图片 | 本地/URL、位置尺寸、四边裁剪、点击超链接 | 已覆盖 | `SlideImageDef`：`crop_left/right/top/bottom`、`hyperlink_url`；SVG 不支持 |
| 形状 | rect/rounded_rect/oval/arrow/cylinder/diamond、填充/边框/虚线/竖排文字/阴影、点击超链接 | 已覆盖 | `SlideShapeDef`；无组合形状/自由曲线/3D/艺术字 |
| 连接线 | 直线连接线、起止箭头、箭头样式、虚线 | 已覆盖 | `SlideConnectorDef` |
| 框架图/流程图 | 阶段 + 节点 + 箭头，17 种节点形状，straight/elbow/curve | 已覆盖 | `ppt_create_framework_diagram`（文件级）、`ppt_add_flowchart_slide`（活动文档） |
| 超链接 | 形状/图片点击链接、指定文字 run 链接 | 已覆盖 | 创建时用 `hyperlink_url`；已有文件用 `ppt_set_hyperlinks`（`shape_index` 或 `shape_text` + `text_match`） |
| 备注 | 演讲者备注写入/读取；`a:br`/`a:fld` 按 python-pptx 语义计入文本 | 已覆盖 | `SlideDef.notes`、`read_powerpoint`；批量 replace 遇到含字段的形状明确拒绝，避免丢字段 |
| 配色 | 12 套预设配色应用到标题/正文/背景/形状/表头 | 已覆盖 | `ppt_list_color_schemes`、`ppt_apply_color_scheme`；只改显式颜色，不改主题定义 |
| 布局审计 | 越界/重叠检查、自动重排、纵向间距优化 | 已覆盖 | `ppt_optimize_slide_layout`、`--reflow`、`--fix-spacing` |
| 读取/分析 | 幻灯片文本/表格/备注、模板尺寸/布局/字体/颜色 | 已覆盖 | `read_powerpoint`、`ppt_analyze_template` |
| 动画 | 进入/强调/退出动画、时间线 | 已覆盖 | `ppt_manage_effects`（office-native COM）：animations-add/list/clear |
| 切换 | 切换效果、时长、自动换片 | 已覆盖 | `ppt_manage_effects`：transition-set/list/clear |
| 批注 | 幻灯片批注 | 已覆盖 | `ppt_manage_effects`：comments-add/list/delete；现代批注作者由 Office 登录身份决定 |
| SmartArt | 创建、列布局、写节点 | 已覆盖 | `ppt_manage_smartart`（office-native COM）：layouts/add；需带窗口打开演示文稿 |
| 母版 | 新增/编辑母版、主题定义 | 未覆盖 | python-pptx 只能改现有 master/layout 形状；新增母版需 XML/COM |
| 页眉页脚/页码 | 页脚文字、幻灯片编号（优先占位符，否则底部文本框 + `a:fld slidenum`） | 已覆盖 | `SlideDef.footer_text`、`show_slide_number` |
| 节/放映设置 | 节分组（`p14:sectionLst`）、放映方式/循环/动画/旁白/计时/页码范围 | 已覆盖 | `PptContent.sections`、`show_settings`（仅创建时）；自定义放映未封装 |
| 媒体/OLE | 音频、视频、嵌入对象 | 未覆盖 | python-pptx 仅实验性 `add_movie`；OLE 只能嵌入 |
| 导出 | PDF、逐页 PNG；单页图片、视频 | 部分 | `convert_to_pdf`、`render_document_pages`；单页图片/视频需 COM（`Slide.Export`、`CreateVideo`） |
| 旧格式 | .ppt / .ppsx | 部分 | `.ppt` 仅转换/COM；文件级编辑只支持 .pptx/.potx |

## 活动文档边界

`edit_open_powerpoint` 只支持基础 paragraph/table/image/shape/connector。本文件新增的表格/图表几何、图片裁剪、形状超链接等字段会**明确报错**（退出码 3），不会静默丢弃；请关闭文件后改用文件级 `edit_powerpoint`。

## 使用建议

标注 office-native 的能力默认走 C#（Windows + PowerPoint）；未构建时入口返回退出码 3 并给出构建命令，或设 `OFFICE_NATIVE_ENGINE=python` 回退旧引擎。

- 做演示文稿：`create_powerpoint`（版式 + 文本 + 表格/图表/图片/形状/连接线）→ `ppt_apply_color_scheme` → `ppt_optimize_slide_layout` → 渲染验收。
- 改已有 PPT：`read_powerpoint` / `inspect_office_document` 看结构 → `batch_edit_office_document` 按形状路径改；链接用 `ppt_set_hyperlinks`。
- 动画、切换、批注、SmartArt：用 `ppt_manage_effects` / `ppt_manage_smartart`（默认 C# COM）；不要用 python-pptx 伪造。
