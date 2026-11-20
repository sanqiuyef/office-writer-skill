# office-native（可选 C# 后端，默认 Python 优先）

office-writer 的 **全量 C# 引擎**：45 个 CLI 入口全部有对应命令，使用 **Open XML SDK 3.5.1** 做文件级读写，**Office COM** 只补独占能力（活动文档、动画/切换/批注/SmartArt、形状绘制、透视表、打印、Office→PDF）。默认执行策略是 Python 优先：只有 `_native.CSHARP_ONLY` 中的 7 个入口（Word 修订/脚注/内容控件、PPT 动画切换批注/SmartArt、Excel 透视表/打印）自动走本后端；其余入口默认走 `scripts/office_engine/`，可用 `OFFICE_NATIVE_ENGINE=csharp` 强制全量 C#。

- Word：创建/编辑/Markdown/读取/校验/模板合并/批注/属性/修订/脚注/内容控件/场景与封面。
- Excel：创建/编辑/读取/8 类图表/条件格式/数据验证/超级表/定义名称/批注/图片/页面设置/透视表/打印。
- PowerPoint：创建/编辑/读取/配色/布局审计与 reflow/清理/模板分析/框架图/超链接/动画/切换/批注/SmartArt。
- PDF：创建（PdfSharpCore，含中文字体）、读取/结构化（PdfPig）、渲染（PDFtoImage/PDFium）。
- 通用：批量路径编辑、转换、节点树、活动文档枚举/替换、纯文本读取。

## 构建

要求：

- .NET SDK 8.0+（本机在 `C:\Program Files\dotnet`）。
- NuGet：`DocumentFormat.OpenXml 3.5.1`、`PdfPig 0.1.11`（Apache-2.0）、`PDFtoImage 5.2.0`（MIT/PDFium）、`PdfSharpCore 1.3.67`（MIT）。
- COM 命令需要 Windows + 对应 Office 应用；SmartArt 需要带窗口打开演示文稿。

```powershell
pwsh -File tools/office-native/build.ps1
# 或
dotnet publish tools/office-native -c Release -r win-x64 --self-contained false `
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
  -o tools/office-native/dist
```

产物：`tools/office-native/dist/office-native.exe`（依赖已安装的 .NET 8 运行时；PDFium/SkiaSharp 原生库已打入单文件并在运行时自解压，约 35 MB）。必须使用 `build.ps1` 或带 `IncludeNativeLibrariesForSelfExtract=true` 的发布命令，否则 OCR/PDF 渲染会报找不到 `pdfium`。

Python 入口按以下顺序查找二进制：

1. `OFFICE_NATIVE_BIN` 环境变量；
2. `tools/office-native/dist/office-native.exe`；
3. `tools/office-native/bin/Release/net8.0/office-native.exe`；
4. `tools/office-native/bin/Debug/net8.0/office-native.exe`。

未构建时入口返回退出码 3，并在错误中给出构建命令。

## 命令

所有命令统一：

```bash
office-native <command> --params '<JSON>' [--params-file <path>]
```

stdout 只输出一个 JSON：`{"ok":true,"tool":"...","data":{...}}`；退出码 0/2/3 与 Python 入口一致。

| 命令 | action | 关键参数 |
|:---|:---|:---|
| `word-revisions` | list / accept-all / reject-all / track-insert / track-delete | `file_path`、`text`、`anchor_text`、`paragraph_index`、`author`、`date` |
| `word-footnotes` | add / list | `file_path`、`text`、`anchor_text`、`paragraph_index` |
| `word-content-controls` | add-text / add-date / add-dropdown / add-checkbox / list / set / remove | `file_path`、`tag`、`alias`、`text`、`date`、`date_format`、`items`、`value`、`checked` |
| `ppt-effects` | animations-add / animations-list / animations-clear / transition-set / transition-clear / transition-list / comments-add / comments-list / comments-delete | `file_path`、`slide`、`shape_index`、`effect`、`duration`、`advance_time`、`text`、`author` |
| `ppt-smartart` | layouts / add | `file_path`、`slide`、`layout`、`texts`、`x/y/width/height`（英寸） |
| `ppt-shapes` | add / list | `file_path`、`slide`、`shape`（`shape_type`、`x/y/width/height` 英寸、`text`、`bg_color`、`font_color`、`font_size`、`bold`、`no_fill`、`line_color`、`line_width`、`line_dash`、`vertical_text`、`hyperlink_url`）；文件已打开时向活动文档添加并保存，未打开时文件级写入并备份 |
| `excel-pivots` | create / list / refresh | `file_path`、`source_sheet`、`source_range`、`target_sheet`、`target_cell`、`name`、`row_field`、`column_field`、`data_field`、`function` |
| `excel-print` | pdf / print | `file_path`、`sheet`、`output`、`printer`、`copies`、`confirm_print` |
| `create-word-document` / `edit-word-document` / `create-word-from-markdown` / `read-word-document` / `validate-word-document` / `word-set-properties` / `word-manage-comments` / `merge-office-template` | — | `output_path`/`file_path`、`content`、`scene`、`cover_config`、`append_sections`、`properties`、`template_path`/`data` |
| `create-excel-workbook` / `edit-excel-workbook` / `read-excel-workbook` / `excel-add-chart` | — | `output_path`/`file_path`、`sheets`/`append_sheets`、`chart`、`max_rows_per_sheet` |
| `create-powerpoint` / `edit-powerpoint` / `read-powerpoint` / `ppt-list-color-schemes` / `ppt-apply-color-scheme` / `ppt-optimize-slide-layout` / `ppt-clean-unused-placeholders` / `ppt-analyze-template` / `ppt-create-framework-diagram` / `ppt-set-hyperlinks` | — | `slides`/`append_slides`、`scheme_name`、`reflow`、`diagram`、`links` |
| `create-pdf-document` / `read-pdf-document` / `read-pdf-structured` / `render-document-pages` | — | `content`、`font_config`、`include_blocks`、`pages`、`dpi` |
| `batch-edit-office-document` / `convert-to-pdf` / `inspect-office-document` / `get-open-office-documents` / `replace-open-document-content` / `read-txt-document` | — | `operations`、`input_path`、`max_nodes`、`replace_sections`/`replace_texts` |
| `edit-open-word-document` / `edit-open-excel-workbook` / `edit-open-powerpoint` / `ppt-add-flowchart-slide` | — | 活动文档 `file_path` + `append_sections`/`append_sheets`/`append_slides`/`diagram` |

## 安全与边界

- **备份**：所有修改型命令在写回前生成 `*.bak-YYYYmmdd-HHMMSS`，结果返回 `backup_path`。
- **覆盖闸门**：`excel-print` 的 PDF 目标已存在且非空时，必须传 `overwrite=true` 与 `confirm_overwrite=true`。
- **COM 进程**：转换/渲染用 `IClassFactory` 创建隔离实例（DispatchEx 语义），不附着用户窗口；活动文档命令只连接已打开实例、不 Quit；所有命令记录本次 PID，只清理自己新起的进程。
- **文件安全**：输入 100 MB 上限、OOXML 解压 500 MB 上限；覆盖先写临时文件、备份后原子替换；备份同秒冲突自动加序号。
- **打印**：`excel-print action=print` 是外部副作用操作，必须显式传 `confirm_print=true`；不指定 printer 时使用默认打印机。
- **SmartArt**：PowerPoint 的 `SmartArtLayouts` 只有在演示文稿带窗口打开时才加载（本机实测无窗口 count=0，带窗口 134）；命令会以带窗口方式打开，处理完关闭。
- **批注作者**：现代 PowerPoint 批注的作者由 Office 登录身份决定，传入的 `author` 可能被忽略；以 `comments-list` 返回为准。
- **Word 修订**：支持 `w:ins` / `w:del` 的插入、删除、接受与拒绝；段落标记修订、移动修订等复杂修订暂不处理。
- **内容控件复选框**：Open XML SDK 3.5 未建模 `w14:checkbox`，用原始 `w14` XML 注入；Word 可正常识别，读取时按本地名/命名空间识别。
- **扫描件 OCR**：PDF 文本层不足时用 PDFium 渲染 + Tesseract CLI；`OFFICE_TESSERACT_CMD`/`TESSERACT_CMD`/PATH 定位，`MCP_OFFICE_OCR_LANG`/`_DPI`/`_MAX_PAGES` 控制；未安装时降级返回文本层。

## 已知边界（非核心路径）

- Excel：形状/绘图、迷你图/切片器、工作簿结构保护、加密、Power Query 仍未封装；富文本单元格 replace 按 Python 语义跳过（保留原格式）；`.xls` 只支持转换。
- Word：段落标记修订、移动修订、邮件合并、比较合并、打印仍未封装；浮动图片/裁剪、自定义样式/主题编辑未开放；水印只写默认页眉。
- PPT：母版、自定义放映、音频视频等媒体跨 slide 移动未封装；组合形状不递归，SVG 不支持；SmartArt 需带窗口的 PowerPoint。
- 活动文档：Excel/PPT 只支持基础追加字段，扩展字段会明确报错并建议改用文件级入口；WPS 个别版本不暴露 `ReadOnly`，由保存失败兜底。
- 图表内嵌工作簿的 part 路径由 SDK 默认生成（`ppt/slides/charts/embeddings/package.xlsx`），PowerPoint 打开/编辑/另存正常，但与 PowerPoint 原生的 `ppt/embeddings/*.xlsx` 路径不保证字节级一致。
- PPT 的 `sections`/`show_settings`/`footer_text`/`show_slide_number` 仅 C# 后端支持；Python 回退引擎会明确报错（不静默忽略）。

## 2026-11-20 新增能力

- Excel：`SheetDef.protection` 工作表保护（密码 + 格式/插入/删除/排序/筛选/选中权限）、`SheetDef.sort` 多键真实排序（写 `sortState`，与合并区冲突报错）。
- Word：`SectionItem.type=bookmark` 书签；`PageSetup.columns`/`column_spacing` 分栏、`page_border`/`page_border_color` 页面边框、`watermark_text`/`watermark_color` 文字水印（默认页眉 VML）。
- PPT：`PptContent.sections` 节（`p14:sectionLst`）、`show_settings` 放映设置（`presProps.xml` 的 `p:showPr`）、`SlideDef.footer_text`/`show_slide_number` 页脚与页码（`a:fld slidenum`，优先占位符）。

## 2026-11-20 补齐的边界

- Excel：日期/时间/时长读取按 openpyxl `data_only` 输出；`max_rows_per_sheet` 流式读取；行列/图表/图片锚点/批注 VML 边界校验；活动 Excel 优先复用打开实例；模板 `=` 值写公式；sharedStrings `extLst` 顺序与计数；表名/范围冲突校验；模板本地名与 activeTab 重算。
- Word：Header/Footer 段落样式；新节强制 `nextPage`；嵌套修订与表格行级修订；跨 run/多 `w:t` 的 track-delete；修订/脚注/内容控件临时文件原子写回；空白新建按 python-docx 默认主题字体。
- 批量：Excel 合并区非锚点写入拒绝；富文本 replace 跳过；区域操作索引化；PPT 跨 slide 移动迁移图片/图表/超链接关系；`if_exists=rename` 真正改名。
- PPT：图表内嵌可编辑数据工作簿（`c:externalData`）；`a:br` 映射 `\v`、`a:fld` 计入文本；含字段形状 replace 明确拒绝。

## 引擎切换

Python 入口默认走 Python 引擎；**`_native.CSHARP_ONLY` 中的 7 个入口**按 `_native.COMMANDS` 映射把同一份参数转给 `office-native.exe`；`OFFICE_NATIVE_ENGINE=csharp`（或 `native`/`1`/`true`）可强制全量 C#，`OFFICE_NATIVE_ENGINE=python`（或 `legacy`/`py`）可强制 Python（C# 专属入口无 Python 实现，仍会调用 C#）。参数通过 stdin 传递，避免 Windows 32K 命令行上限。

```bash
python scripts/read_txt_document.py --file-path a.txt
OFFICE_NATIVE_ENGINE=python python scripts/read_txt_document.py --file-path a.txt
OFFICE_CONVERT_BACKEND=com python scripts/convert_to_pdf.py --input-path a.xlsx --output-path a.pdf
```

已完成 C# 实现的入口：全部 45 个（Word/Excel/PPT 文件级、PDF、通用、活动文档、批量编辑、PPT 形状绘制）。`tests/test_native_parity.py` 对 Word/Excel/PPT 创建、Word 校验、txt 读取做 Python↔C# 等价对比。

## 与 Python 入口的关系

45 个 Python 脚本只做参数校验与 JSON 转发；能力实现全部在本后端。旧 Python 引擎（`scripts/office_engine/`）保留为 `OFFICE_NATIVE_ENGINE=python` 的对照实现。
