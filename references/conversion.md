# PDF 转换、页面渲染与公式重算（dsoffice / Office COM / WPS COM）

来源：mcp-office-writer 0.6.0 的 `office_convert.py`（2026-11）、`@deepseek-ai/libreoffice-kit` 的 dsoffice CLI 行为，以及 2026-11-15 本机 Office 16.0 / WPS 12.1 COM 实测。

适用边界：需要 Office → PDF、逐页 PNG、xlsx 公式缓存值重算时使用。不覆盖 PDF 编辑、拆分、合并、加密；COM 回退仅 Windows。

## 后端顺序

`OFFICE_CONVERT_BACKEND` 控制三个入口（`convert_to_pdf`、`render_document_pages`、`recalculate_excel_formulas`）：

| 值 | 行为 |
|:---|:---|
| `auto`（默认） | dsoffice 可用则先用；缺失或执行失败回退 Microsoft Office COM，再回退 WPS COM |
| `dsoffice` | 只用 dsoffice；缺失/失败直接报错，不回退 |
| `com` | 只用 COM：Office ProgID 优先，WPS `KWps/KET/KWPP` 次之 |
| `none` | 禁用 Office 转换/COM 渲染/重算；直接对 `.pdf` 输入渲染仍可用 PyMuPDF。错误信息列出三种启用方式 |

`auto` 回退成功时结果带 `fallback_from[]`，例如：

```json
"fallback_from": [{"backend": "dsoffice", "error": "MODULE_OFFICE_DSOFFICE 指定的路径不存在: ..."}]
```

## dsoffice

```bash
npm i -g dsoffice
# 要求 Node.js >= 22.19；Windows x64 另需 MSVC v14 运行库
which dsoffice     # PowerShell: Get-Command dsoffice
```

定位顺序：

1. `MCP_OFFICE_DSOFFICE` 指定的可执行文件（兼容 `MODULE_OFFICE_DSOFFICE`，前者优先）；
2. `PATH` 中的 `dsoffice`。

首次运行可能下载引擎，超时上限 300 秒。dsoffice 输出目标必须不存在：入口先写同目录临时文件，成功后再原子替换，失败不破坏已有目标。

## Microsoft Office / WPS COM

按扩展名选应用与 ProgID：

| 格式 | Office | WPS 优先 | WPS 再回退 |
|:---|:---|:---|:---|
| docx/doc/rtf | `Word.Application` | `KWps.Application` | `Word.Application` |
| xlsx/xls | `Excel.Application` | `KET.Application` | `Excel.Application` |
| pptx/ppt | `PowerPoint.Application` | `KWPP.Application` | `PowerPoint.Application` |

实现要点：

- `win32com.client.DispatchEx(progid)` 启动独立隐藏实例，不复用用户已打开的窗口；
- Word：`Documents.Open(..., ReadOnly=True)` → `ExportAsFixedFormat(pdf, 17)`；
- Excel：`Workbooks.Open(..., ReadOnly=True)` → `Workbook.ExportAsFixedFormat(0, pdf)`；指定 `sheet` 时用 `Worksheet.ExportAsFixedFormat`，`cell_range` 写成 `PageSetup.PrintArea`；
- PowerPoint：`Presentations.Open(..., WithWindow=False)` → `SaveAs(pdf, 32)`；
- 无论成功失败都关闭文档并 `Quit`；拿到 PID 时等待退出，必要时 `taskkill /PID ... /F /T`；
- 结果返回 `backend=com`、`backend_kind=com`、`com_engine=office|wps`、`com_prog_id`、`com_application`。

COM 要求 Windows + pywin32 + 已安装 Office/WPS。WPS 常注册与 Office 同名的 ProgID，因此 `com_engine` 由实际应用路径判定，不能只看 ProgID。

## convert_to_pdf

```bash
python -X utf8 scripts/convert_to_pdf.py \
  --input-path 报告.docx \
  --output-path 报告.pdf \
  --if-exists overwrite
```

- 输入：`.docx/.doc/.xlsx/.xls/.pptx/.ppt/.rtf`；输出 `.pdf`；
- 省略 `--output-path` 时写源文件同目录同名 `.pdf`；
- Excel 可用 `--sheet` 只转换一个工作表（dsoffice 与 Excel COM 都支持）；
- `if_exists`：`overwrite` / `error` / `rename`；
- 返回 `file_path`、`backend`、`backend_kind`、`com_engine`、`com_prog_id`、`missing_fonts[]`、`font_substitutions[]`、`message`；
- `font_substitutions` 只覆盖等宽字体（Consolas → Cascadia Mono → DejaVu Sans Mono），找不到替代时为空。

## render_document_pages

```bash
python -X utf8 scripts/render_document_pages.py \
  --input-path 报告.pdf \
  --output-dir 报告_pages \
  --pages 1,3 \
  --dpi 144
```

- 输入：convert 支持的全部格式 + `.pdf`；
- 输出目录默认 `<源文件名>_pages`；`if_exists=overwrite/error`；
- `--pages all|1,3,5`；`--dpi` 默认 144，夹到 72–600；
- Excel 的 `--sheet` + `--cell-range A1:D20` 在 dsoffice 下按工作表区域渲染；COM 回退把区域写成 print area，再渲染整页；
- dsoffice 缺失或失败时：PDF 直接由 PyMuPDF 渲染；Office 文档先经 COM 转临时 PDF，再逐页出 PNG；
- 返回 `output_dir`、`page_count`（实际 PNG 数）、`images[]`（`index/page/width/height/path`）、`backend`、`backend_kind`、`fallback_from[]`；
- `page_count` 以输出目录实际 PNG 数为准；与 PDF 总页数不一致时返回 `page_count_reported`。

## recalculate_excel_formulas

```bash
python -X utf8 scripts/recalculate_excel_formulas.py \
  --input-path 公式.xlsx \
  --output-path 公式_recalc.xlsx
```

- 输入 `.xlsx/.ods`，输出 `.xlsx/.ods`；省略输出时默认 `<原名>_recalc.<原后缀>`；
- dsoffice：整体重算后写回缓存值；
- COM 回退：复制到临时 `.xlsx` → `Application.CalculateFullRebuild()`（失败时 `Calculate()`）→ `Workbook.Save()` → 替换目标；
- COM 回退只支持 `.xlsx`；`.ods` 需安装 dsoffice；
- 返回 `backend`、`backend_kind`、`com_prog_id`、`calculation`、`missing_fonts[]`。

## 失败形态

| 现象 | 原因 | 处理 |
|:---|:---|:---|
| `未找到可用的 Office 转换后端` | dsoffice 缺失且 COM 不可用 | 按错误里的三种方式安装/启用其一 |
| `MCP_OFFICE_DSOFFICE 指定的路径不存在` | 环境变量指向错误路径 | 修正变量，或设为空让 PATH 查找 |
| `dsoffice 执行超时` | 首次下载引擎/大文件 | 重试；检查网络；拆分大文件 |
| `dsoffice 输出无法解析` | CLI 版本不兼容或崩溃 | 更新 `@deepseek-ai/libreoffice-kit`，查看 stderr |
| `COM 转换失败（已尝试 ...）` | Office/WPS 未装、弹窗、文件损坏 | 看各 ProgID 错误；安装/修复 Office 或 WPS；改用 dsoffice |
| `文件已存在` | `if_exists=error` | 改用 `overwrite`/`rename` |
| `pages 参数无效` | 页码格式错误 | 用 `all` 或 `1,3,5` |
| `Excel 渲染需要同时指定 sheet 和 cell_range` | 只给了区域或缺工作表 | 同时传 `--sheet` 与 `--cell-range` |
| `missing_fonts` 非空 | 源文档字体缺失 | 不阻断转换；确认版式或改用通用字体 |

## 实测记录（2026-11-15）

- dsoffice：报告 docx、数据 xlsx、演示 pptx 均转出 PDF（66401 / 27224 / 45524 字节）；docx 报 `Courier` 缺失并建议 `Consolas`。
- COM：docx→`Word.Application`、xlsx→`Excel.Application`、pptx→`PowerPoint.Application` 均退出 0，`com_engine=office`；WPS `KWps/KET/KWPP` 也曾实际完成导出。
- 渲染回退：`MODULE_OFFICE_DSOFFICE` 指向不存在路径时，A4 docx 经 COM + PyMuPDF 在 144dpi 生成 1191×1684 PNG。
- 公式重算：COM `CalculateFullRebuild` 后回读合计列 `255/390`。
- 进程清理：COM 调用结束后 `tasklist` 无残留 `WINWORD/EXCEL/POWERPNT`；Word 进程退出可能滞后数秒。

## OCR（可选）

`read_pdf_document.py` 发现 PDF 文本层缺失/几乎只有重复水印时，会尝试 Tesseract OCR（Python 与 C# 后端行为一致）：

- Python 后端需要 `pytesseract` + Tesseract 可执行文件；C# 后端直接调用 Tesseract CLI；
- 可执行文件定位：`OFFICE_TESSERACT_CMD`，其次 `TESSERACT_CMD`，最后从 `PATH` 查找 `tesseract`；
- 开关与参数：`MCP_OFFICE_OCR`、`MCP_OFFICE_OCR_LANG`、`MCP_OFFICE_OCR_DPI`、`MCP_OFFICE_OCR_MAX_PAGES`、`MCP_OFFICE_OCR_MIN_TEXT_CHARS`；
- 未安装 Tesseract 时，文本 PDF 不受影响；扫描件读取降级返回文本层并在 `message` 中说明。
