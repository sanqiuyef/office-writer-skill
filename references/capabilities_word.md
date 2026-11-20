# Word 能力矩阵

来源：Microsoft Word VBA 对象模型（learn.microsoft.com/office/vba/api/word.*）、python-docx 1.2.0 API、本 skill 的 C# office-native 引擎实现（2026-11）。
适用边界：Windows 桌面 Word 的 COM 自动化与文件级 Open XML 读写；不含 Word 网页版、macOS、VBA 宏执行、邮件合并。

状态记号：**已覆盖** = 有入口且已实测；**部分** = 能力有限制；**未覆盖** = 当前没有入口，括号内给替代路径。

| 功能域 | 能力项 | 状态 | 入口 / 参数 / 限制 |
|:---|:---|:---|:---|
| 文档生命周期 | 新建 .docx、模板继承、场景排版、封面 | 已覆盖 | `create_word_document`：`--template`、`--scene`、`cover_config` |
| 文档生命周期 | 读取正文/表格、质量门控 | 已覆盖 | `read_word_document`、`validate_word_document` |
| 文档结构 | 标题、段落、分页、分节（新节强制从新页开始）、项目符号/编号列表、代码块 | 已覆盖 | `SectionItem.type`；编号用内置 List Number 样式，无自定义多级编号 |
| 样式与字体 | 中文字号、字体、粗斜体、颜色、对齐、行距、缩进、段前段后；空白新建按 python-docx 默认主题（正文 Cambria/宋体，标题 Calibri/宋体） | 已覆盖 | `ParagraphStyle`、`default_font`；无自定义样式/主题编辑 |
| 表格 | 表头样式、列宽、A1 范围合并 | 已覆盖 | `TableDef`；无单元格边框/底纹/行高/重复表头 |
| 图片 | 本地图片、指定宽高 | 已覆盖 | `ImageDef`；仅 inline，无浮动/环绕/裁剪/Alt Text |
| 页眉页脚 | 文本页眉、页脚、`{page}` 页码域；新建页眉/页脚段落使用 Header/Footer 样式（0 段后、单倍行距、左右 tab） | 已覆盖 | `HeaderFooter`；无首页/奇偶页差异，无 NUMPAGES 自动文本（用 field） |
| 域 | PAGE / NUMPAGES / SECTIONPAGES / DATE / TIME / FILENAME / AUTHOR / TITLE / CUSTOM | 已覆盖 | `SectionItem.type=field` + `FieldDef`；CUSTOM 用 `code` 传完整域代码 |
| 目录 | TOC 域（1-9 级） | 已覆盖 | `SectionItem.type=toc`；需在 Word 中更新域后显示条目 |
| 超链接 | 外部 URL 超链接、tooltip | 已覆盖 | `SectionItem.type=hyperlink` + `HyperlinkDef` |
| 批注 | 新增批注、列出批注（作者/时间/锚定文字） | 已覆盖 | `word_manage_comments`：`anchor_text` 或 `paragraph_index` 定位；无回复/解决/删除 |
| 文档属性 | 标题/主题/作者/关键词/备注/类别/最后修改者/创建与修改时间 | 已覆盖 | `WordContent.properties`（创建时）、`word_set_properties`（已有文件）；无自定义属性 |
| 查找替换 | 文件级逐段替换、活动文档全文替换 | 已覆盖 | `batch_edit_office_document`（replace）、`replace_open_document_content` |
| 模板填充 | `{{key}}` 占位符、嵌套对象/数组 | 已覆盖 | `merge_office_template`；跨 run 标签无法保证 |
| Markdown | 长文 Markdown → Word | 已覆盖 | `create_word_from_markdown` |
| 转换/验收 | docx → PDF、逐页 PNG、结构检查 | 已覆盖 | `convert_to_pdf`、`render_document_pages`、`validate_word_document` |
| 修订 | 列出、接受/拒绝、标记插入与删除；支持嵌套 `ins/del`、表格行级修订、跨 run/多 `w:t` 的删除 | 已覆盖 | `word_manage_revisions`（office-native Open XML SDK）；段落标记修订、移动修订未处理 |
| 脚注 | 插入、读取、编号 | 已覆盖 | `word_manage_footnotes`（office-native）；尾注未封装 |
| 书签 | 插入书签、按书签定位 | 未覆盖 | 需要原始 OOXML 或 COM |
| 内容控件 | 文本/日期/下拉/复选框的增删改查 | 已覆盖 | `word_manage_content_controls`（office-native）；构建基块未封装 |
| 邮件合并 | 数据源、域映射、批量生成 | 未覆盖 | 需要 COM（MailMerge） |
| 文档比较/合并 | CompareDocuments、MergeDocuments | 未覆盖 | 需要 COM |
| 打印 | PrintOut、打印到指定打印机 | 未覆盖 | 需要 COM；渲染/PDF 可作替代 |
| 分栏/页面边框/水印 | 多栏、页面边框、水印 | 未覆盖 | 需要原始 OOXML 或页眉 Shape 模拟（水印） |
| 图表/SmartArt/OLE/公式 | 创建与编辑 | 未覆盖 | 需要 COM 或 Open XML SDK；打开/保存可保留已有对象 |
| 旧格式 | .doc / .dotm / .docm | 不支持 | 文件级读写仅 .docx（模板也按 .docx 处理） |

## 活动文档边界

`edit_open_word_document` 只支持 heading/paragraph/bullet_list/numbered_list/page_break；域、超链接、TOC、分节、表格、图片等请关闭文件后走文件级 `edit_word_document` 或 `batch_edit_office_document`。

## 使用建议

标注 office-native 的能力默认走 C#（Windows + Office）；未构建时入口返回退出码 3 并给出构建命令，或设 `OFFICE_NATIVE_ENGINE=python` 回退旧引擎。

- 生成新文档：`create_word_document`（内容块 + 场景 + 属性）→ `validate_word_document`。
- 改已有文档：优先 `edit_word_document` 追加，或 `batch_edit_office_document` 按路径改；不要用 `create_*` 覆盖。
- 审阅：`word_manage_comments` 增/列批注；修订、脚注、内容控件用 `word_manage_*`（默认 C#）。
