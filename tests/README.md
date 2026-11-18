# tests/

本目录记录三类验证：确定性 CLI 契约、结构/触发基线、预存 golden 的 parity。真实文档往返与转换/渲染记录在最后一节。

## 契约测试

```bash
cd <本项目目录>
python -X utf8 -m unittest discover -s tests -v
```

本轮结果：**42 项通过，0 跳过**。覆盖：

- 11 个固定用例的 `data` 全量比对（读取、结构化读取、inspect、validate、模板分析、配色列表）；
- 32 个入口逐个 `--help`，全部退出 0；
- 退出码 2：缺参数、非法 JSON、非法枚举、Pydantic 校验失败；
- 退出码 3：文件不存在、转换输入缺失、PPT 配色目标缺失、缺 dsoffice 且错误含安装/检测提示、`OFFICE_CONVERT_BACKEND=none` 列出三种启用方式、非法后端值；
- 创建 → 读取 → validate → inspect 往返（docx/xlsx/pptx/pdf）；
- Markdown 转 Word、追加段落、Word/Excel/PPT 模板合并、批量编辑、PPT 配色/布局/清理、框架图；
- 渲染 `page_count` 与实际 PNG 数一致（fake dsoffice 报告 2 页、实际 1 张时修正为 1，并返回 `page_count_reported=2`）；
- 字体回退：`plan_font_substitutions` 的候选顺序、无候选时为空，以及 fake dsoffice 的 `font_substitutions` 输出；
- 批量原地编辑生成 `*.bak-YYYYmmdd-HHMMSS`，备份内容等于原文件；`--no-backup` 与显式 `--output-path` 不生成备份；
- PPT 审计识别重叠与越界、默认不改文件，`--fix-spacing` 仍写回；
- PPT `--reflow`：越界、重叠、混合、超大四类几何重排后无剩余 warning，文本/字体/颜色不变；
- validate 实检：干净 fixture 17/17 通过；问题 fixture 命中字体、标题层级/编号、PAGE、TOC、表格列、空段落、中英混排；TOC 未标脏 fixture 单独命中；
- COM 真实路径：`OFFICE_CONVERT_BACKEND=com` 转换 docx/xlsx/pptx、重算 xlsx，断言 `backend=com`、PDF 头、缓存值，并在 20 秒内确认无残留 `WINWORD/EXCEL/POWERPNT`；
- 缺 dsoffice 回退：`MODULE_OFFICE_DSOFFICE` 指向不存在路径时，docx 经 COM 转 PDF；docx 渲染经 COM+PyMuPDF 出 PNG，`page_count` 与 PNG 数一致；
- Word 域 XML 兼容：生成的 report/TOC 文档中 `fldChar`/`instrText` 均包在 `w:r` 中，`pBdr` 位于 `spacing/ind/jc` 之前；
- 扫描件 OCR：Pillow 生成图片 PDF，经 `OFFICE_TESSERACT_CMD` 调用真实 Tesseract；
- stdout 单 JSON 对象、`--params-file`。

有意改变固定 CLI 行为后重新生成期望值（parity golden 保持静态）：

```bash
python -X utf8 tests/update_expected.py
```

## 触发基线

`test_triggers.py` 从 `SKILL.md` 的 `触发：` 词表做词法匹配，对 `trigger_cases.json` 的 21 条正例、10 条负例断言 100% 一致；正例的 `expected_script` 必须存在。它验证的是触发词与路由基线，不替代宿主模型的实际路由评测。

## parity（golden，不依赖 mcp）

`test_parity.py` 读取 `tests/expected_outputs/parity_*.json`，与 vendored 引擎的当前输出逐字段比较：

1. `ppt_writer.list_color_schemes()`；
2. `reader.read_txt_document()` 读取 `tests/fixtures/sample.txt`；
3. `checker.validate_word_document().to_dict()` 读取 `tests/fixtures/sample.docx`。

三个 golden 的原始版本于 2026-11-15 与 mcp-office-writer 0.6.0 的对应输出逐字段核对一致。本轮有意给 `validate_word_document` 增加 7 项确定性检查和稳定 `id`，因此 `parity_validate_word.json` 已按当前引擎重新记录（17 项检查）；`parity_list_color_schemes.json` 与 `parity_read_txt.json` 保持上游快照。`update_expected.py` 只刷新固定 CLI 用例，不会自动改写 parity golden，避免覆盖上游差异。测试文件不再 `import mcp_office_writer`，结构测试的零 MCP 依赖检查同时扫描 `scripts/` 与 `tests/`。

## 真实 CLI 验证（2026-11-15）

工作目录：临时目录。解释器：Python ≥ 3.10。

### 1. 创建四类文档

本轮实际执行的完整命令（工作目录 `scripts/`）：

```bash
python -X utf8 create_word_document.py --output-path .../report.docx --params '{"scene":"report","content":{"sections":[{"type":"heading","text":"2026 年第三季度报告","level":1},{"type":"paragraph","text":"本报告汇总项目进展、关键指标与下一步计划。"},{"type":"heading","text":"一、关键指标","level":2},{"type":"table","table":{"headers":["指标","目标","实际"],"rows":[["交付率","95%","97%"],["缺陷率","2%","1.4%"]]}},{"type":"bullet_list","items":["完成集成测试","补充用户文档","启动验收评审"]}]}}'

python -X utf8 create_excel_workbook.py --output-path .../data.xlsx --params '{"sheets":[{"name":"数据","headers":["区域","一季度","二季度","合计"],"data":[["华北",120,135,"=B2+C2"],["华东",180,210,"=B3+C3"],["华南",150,165,"=B4+C4"]],"allow_formulas":true,"freeze":"A2","auto_filter":true,"column_widths":[12,10,10,10]},{"name":"汇总","headers":["说明","值"],"data":[["周期","2026 Q1-Q2"],["来源","销售系统"]]}]}'

python -X utf8 create_powerpoint.py --output-path .../deck.pptx --params '{"slides":[{"layout":"title","title":"2026 年第三季度汇报","subtitle":"项目办公室"},{"layout":"content","title":"关键进展","body":[{"type":"bullet_list","items":["完成核心模块开发","交付率 97%","缺陷率降至 1.4%"]}]},{"layout":"title_only","title":"指标表","body":[{"type":"table","table":{"headers":["指标","目标","实际"],"rows":[["交付率","95%","97%"],["缺陷率","2%","1.4%"]]}},{"type":"shape","shape":{"shape_type":"rounded_rect","text":"下一步：验收","x":5.5,"y":5.5,"width":3.2,"height":0.8,"bg_color":"2563EB","font_color":"FFFFFF"}}]}]}'

python -X utf8 create_pdf_document.py --output-path .../document.pdf --params '{"content":{"sections":[{"type":"heading","text":"2026 年第三季度报告","level":1},{"type":"paragraph","text":"本报告汇总项目进展、关键指标与下一步计划。"},{"type":"heading","text":"关键指标","level":2},{"type":"table","table":{"headers":["指标","目标","实际"],"rows":[["交付率","95%","97%"],["缺陷率","2%","1.4%"]]}},{"type":"bullet_list","items":["完成集成测试","补充用户文档","启动验收评审"]}]}}'
```

输出摘要（均为退出 0）：

| 文件 | data.success | file_size | 其他 |
|:---|:---|:---|:---|
| `report.docx` | true | 37893 | 场景 report，含封面 |
| `data.xlsx` | true | 5772 | 2 个工作表，3 个公式 |
| `deck.pptx` | true | 30803 | 3 张幻灯片 |
| `document.pdf` | true | 23276 | 5 个内容块 |

### 2. 回读 + validate + inspect

```bash
python -X utf8 scripts/read_word_document.py --file-path ...\report.docx
python -X utf8 scripts/read_excel_workbook.py --file-path ...\data.xlsx
python -X utf8 scripts/read_powerpoint.py --file-path ...\deck.pptx
python -X utf8 scripts/read_pdf_document.py --file-path ...\document.pdf
python -X utf8 scripts/validate_word_document.py --file-path ...\report.docx
python -X utf8 scripts/inspect_office_document.py --file-path ...\report.docx
python -X utf8 scripts/inspect_office_document.py --file-path ...\data.xlsx
python -X utf8 scripts/inspect_office_document.py --file-path ...\deck.pptx
```

结果（均退出 0）：

- Word：`成功读取文档，共 7 段落、1 个表格`；markdown 224 字符。
- Excel：`成功读取工作簿，共 2 个工作表、7 行数据`；公式列暂为空（openpyxl 无缓存值）。
- PPT：`成功读取演示文稿，共 3 张幻灯片`；markdown 279 字符。
- PDF：`成功读取 PDF，共 1 页`；markdown 128 字符。
- validate：`total=10, passed=10, warnings=0, errors=0`，退出 0。
- inspect：docx 26 节点 / xlsx 25 节点 / pptx 12 节点，`truncated=false`。

### 3. 公式重算

```bash
python -X utf8 scripts/recalculate_excel_formulas.py --input-path ...\data.xlsx --output-path ...\data_recalc.xlsx
python -X utf8 scripts/read_excel_workbook.py --file-path ...\data_recalc.xlsx
```

结果：退出 0，`backend=native`，`missing_fonts=[]`；回读合计列包含 `255 / 390 / 315`。

### 4. 转 PDF 与字体回退

```bash
python -X utf8 scripts/convert_to_pdf.py --input-path ...\report.docx --output-path ...\report.pdf
python -X utf8 scripts/convert_to_pdf.py --input-path ...\data.xlsx --output-path ...\data.pdf
python -X utf8 scripts/convert_to_pdf.py --input-path ...\deck.pptx --output-path ...\deck.pdf
```

结果（均退出 0，`backend=native`）：

- `report.pdf` 81854 字节，`missing_fonts=["Courier"]`，`font_substitutions=[{"missing":"Courier","substitute":"Consolas"}]`；
- `data.pdf` 31353 字节，无缺失字体；
- `deck.pdf` 56933 字节，无缺失字体。

### 5. 渲染页面与计数修正

```bash
python -X utf8 scripts/render_document_pages.py --input-path ...\report.docx --output-dir ...\report_docx_pages --pages 1
python -X utf8 scripts/render_document_pages.py --input-path ...\report.pdf --output-dir ...\report_pdf_pages --pages all
python -X utf8 scripts/render_document_pages.py --input-path ...\document.pdf --output-dir ...\document_pages --pages all
python -X utf8 scripts/render_document_pages.py --input-path ...\data.xlsx --output-dir ...\data_pages --sheet 数据 --cell-range A1:D4
python -X utf8 scripts/render_document_pages.py --input-path ...\deck.pdf --output-dir ...\deck_pages --pages all
```

结果（均退出 0）：

| 输入 | page_count | page_count_reported | images | 实际 PNG |
|:---|:---|:---|:---|:---|
| `report.docx --pages 1` | 1 | 2 | 1 | 1 |
| `report.pdf --pages all` | 2 | 无 | 2 | 2 |
| `document.pdf --pages all` | 1 | 无 | 1 | 1 |
| `data.xlsx --sheet 数据 --cell-range A1:D4` | 1 | 2 | 1 | 1 |
| `deck.pdf --pages all` | 3 | 无 | 3 | 3 |

`report.docx` 与 `data.xlsx` 的上游元数据都报告 2，入口已按实际产物修正为 1；`page_count_reported` 保留原值。

### 6. 扫描件 OCR

```bash
# Pillow 生成白底图片 -> PDF（无文本层）
python - <<'PY'
from PIL import Image, ImageDraw, ImageFont
image = Image.new("RGB", (1600, 600), "white")
draw = ImageDraw.Draw(image)
draw.text((80, 220), "SCANNED INVOICE 2026", fill="black", font=ImageFont.truetype("arial.ttf", 96))
image.save("scan.pdf", "PDF", resolution=150.0)
PY

OFFICE_TESSERACT_CMD="<tesseract 可执行文件路径>" \
MCP_OFFICE_OCR_LANG=eng \
python -X utf8 scripts/read_pdf_document.py --file-path ...\scan.pdf
```

结果：退出 0，`message=成功读取 PDF（OCR），共 1 页`，识别文本包含 `SCANNED INVOICE 2026`。

### 7. COM Word 成功路径

用 `win32com.client.DispatchEx("Word.Application")` 打开临时 `com_probe.docx`（`Visible=False`、`DisplayAlerts=0`），再调用：

```bash
python -X utf8 scripts/get_open_office_documents.py
python -X utf8 scripts/edit_open_word_document.py --file-path ...\com_probe.docx --params '{"append_sections":[{"type":"paragraph","text":"COM 成功路径实测追加段落 2026-11-15。"}]}'
python -X utf8 scripts/read_word_document.py --file-path ...\com_probe.docx
```

结果：

- `get_open_office_documents` 退出 0，`word` 列表 1 项，`application=Word.Application`；
- `edit_open_word_document` 退出 0，`operation_count=1`，`message=已通过 Word.Application 修改并保存活动 Word 文档`；
- COM 内 `doc.Content.Text` 与回读 markdown 均含追加段落，`doc.Saved=True`；
- `doc.Close(0)` + `app.Quit(0)` 后等待 3 秒复查，无残留 `WINWORD` 进程。

### 8. 批量原地编辑备份

```bash
python -X utf8 scripts/batch_edit_office_document.py --file-path ...\batch_inplace.docx \
  --params '{"operations":[{"op":"set","path":"/body/p[2]","props":{"text":"批量编辑已原地写回。"}}]}'
python -X utf8 scripts/batch_edit_office_document.py --file-path ...\batch_inplace.docx --no-backup \
  --params '{"operations":[{"op":"set","path":"/body/p[2]","props":{"text":"第二次原地写回，无备份。"}}]}'
```

结果：第一次退出 0，生成 `batch_inplace.docx.bak-20261115-122524`（37893 字节），`backup_path` 指向该文件；第二次退出 0，结果无 `backup_path`，备份仍只有 1 个。

### 9. PPT 越界/重叠审计

构造 3 个形状（A/B 重叠，C 越右下边界）后：

```bash
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path ...\layout_test.pptx
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path ...\layout_test.pptx --fix-spacing
```

结果：第一次退出 0，`slide_count=1`、`shape_count=3`、`warning_count=2`：

- `out_of_bounds`：`#3 Rectangle 3` 右边界超出 2.00in、下边界超出 0.50in；
- `overlap`：`#1 Rectangle 1` 与 `#2 Rectangle 2` 重叠约 2.00×1.00in。

第二次退出 0，`slides_optimized=1`，文件被写回。契约测试另外断言默认审计不修改文件。

### 10. 缺 dsoffice 的错误提示（强制 dsoffice）

```bash
OFFICE_CONVERT_BACKEND=dsoffice MCP_OFFICE_DSOFFICE= MODULE_OFFICE_DSOFFICE= PATH= \
  python -X utf8 scripts/convert_to_pdf.py --input-path ...\dsoffice_missing.docx --output-path ...\dsoffice_missing.pdf
```

结果：退出 3，`error` 含 `npm i -g dsoffice`、`Node.js >= 22.19`、`Get-Command dsoffice`；目标 PDF 未被创建。查找 dsoffice 发生在替换旧输出之前。默认 `auto` 在同样条件下会回退 COM，不再报错（见 11.2）。

## 本轮增量验收（2026-11-15 第二轮）

工作目录：临时目录。解释器：Python ≥ 3.10。

### 11.1 转换：dsoffice 与 COM 各跑 docx/xlsx/pptx

```bash
# dsoffice（默认 auto 且 PATH 中有 dsoffice）
python -X utf8 scripts/convert_to_pdf.py --input-path ...\report_final.docx --output-path ...\dsoffice_docx.pdf --if-exists overwrite
python -X utf8 scripts/convert_to_pdf.py --input-path ...\data.xlsx --output-path ...\dsoffice_xlsx.pdf --if-exists overwrite
python -X utf8 scripts/convert_to_pdf.py --input-path ...\deck.pptx --output-path ...\dsoffice_pptx.pdf --if-exists overwrite

# COM
OFFICE_CONVERT_BACKEND=com python -X utf8 scripts/convert_to_pdf.py --input-path ...\report_final.docx --output-path ...\com_docx.pdf --if-exists overwrite
OFFICE_CONVERT_BACKEND=com python -X utf8 scripts/convert_to_pdf.py --input-path ...\data.xlsx --output-path ...\com_xlsx.pdf --if-exists overwrite
OFFICE_CONVERT_BACKEND=com python -X utf8 scripts/convert_to_pdf.py --input-path ...\deck.pptx --output-path ...\com_pptx.pdf --if-exists overwrite
```

结果（均退出 0）：

| 路径 | backend | com_prog_id | 字节 |
|:---|:---|:---|:---|
| `dsoffice_docx.pdf` | native / dsoffice | - | 66401，`missing_fonts=["Courier"]` |
| `dsoffice_xlsx.pdf` | native / dsoffice | - | 27224 |
| `dsoffice_pptx.pdf` | native / dsoffice | - | 45524 |
| `com_docx.pdf` | com / office | `Word.Application` | 140417 |
| `com_xlsx.pdf` | com / office | `Excel.Application` | 18859 |
| `com_pptx.pdf` | com / office | `PowerPoint.Application` | 54503 |

调试过程中也实际走到过 WPS 分支：`com_engine=wps`、`com_prog_id=KWps.Application`，说明 Office 优先、WPS 兜底的链有效。

### 11.2 缺 dsoffice：COM 转换与渲染回退

```bash
MODULE_OFFICE_DSOFFICE=...\missing-dsoffice.exe \
  python -X utf8 scripts/render_document_pages.py \
  --input-path ...\report_final.docx --output-dir ...\render_fallback_pages \
  --pages 1 --dpi 144 --if-exists overwrite
```

结果：退出 0，`backend=com+pymupdf`、`backend_kind=com`、`com_prog_id=Word.Application`、`page_count=1`；生成 `page-0001.png`（1191×1684），目录内没有临时 PDF。`fallback_from=[{"backend":"dsoffice","error":"MODULE_OFFICE_DSOFFICE 指定的路径不存在: ..."}]`。

### 11.3 COM 公式重算

```bash
OFFICE_CONVERT_BACKEND=com python -X utf8 scripts/recalculate_excel_formulas.py \
  --input-path ...\data.xlsx --output-path ...\data_recalc.xlsx --if-exists overwrite
```

结果：退出 0，`backend=com`、`com_prog_id=Excel.Application`、`calculation=CalculateFullRebuild`；openpyxl `data_only=True` 回读合计列 `255 / 390`。

### 11.4 validate 新 fixture

fixture：`tests/fixtures/validate_clean.docx`、`validate_issues.docx`、`validate_toc_stale.docx`。

| fixture | total | passed | warnings | errors | 命中 |
|:---|---:|---:|---:|---:|:---|
| `validate_clean.docx` | 17 | 17 | 0 | 0 | 无 |
| `validate_issues.docx` | 17 | 9 | 8 | 0 | 标题层级/编号、字体未装、PAGE、TOC、表格列、空段落、中英混排 |
| `validate_toc_stale.docx` | 17 | 16 | 1 | 0 | 目录域存在但未标记为需更新 |

`validate_issues.docx` 的具体 warning：`font_installed`（Courier、不存在的字体）、`heading_levels`（H1→H3）、`heading_numbering`（1→3）、`page_number_field`、`toc_field`、`table_columns`（第 2 行 2 列 vs 首行 3 列）、`empty_paragraph_ratio`（3/10，最长连续 3）、`mixed_fonts`（宋体/黑体 + 中文 run 只设 Arial）。

### 11.5 PPT reflow 前后坐标

```bash
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path ...\mixed.pptx
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path ...\mixed.pptx --reflow
python -X utf8 scripts/ppt_optimize_slide_layout.py --file-path ...\mixed.pptx
```

| 用例 | 重排前 warning | reflow 记录（from → to） | 重排后 |
|:---|:---|:---|:---|
| 越界 | 右 2.00in、下 0.50in | `#2 move`: (9.0,6.0,3.0,2.0) → (6.9,5.4,3.0,2.0) | 0 warning |
| 重叠 | 2.00×1.00in | `#3 move`: (3.0,2.0,4.0,2.0) → (3.0,3.15,4.0,2.0) | 0 warning |
| 混合 | 越界 + 重叠 | `#3 move` + `#4 move`（同上） | 0 warning |
| 超大 | 右 2.50in、下 3.00in | `#2 move_resize`: (0.5,1.5,12×9) → (0.167,0.1,9.733×7.3) | 0 warning |

### 11.6 Word XML 兼容修正

真实 COM 调试发现：`Word.Application` 拒绝打开本 skill 生成的封面报告，错误为“Word 在试图打开文件时遇到错误”；WPS/LibreOffice 仍能打开。根因是两处 OOXML 顺序/结构问题：

1. 页脚 PAGE 域与 TOC 域的 `w:fldChar`/`w:instrText` 直接挂在 `w:p` 下，没有包在 `w:r` 中；
2. 封面水平线的 `w:pBdr` 被追加在 `w:spacing`/`w:jc` 之后。

修正 `word_writer.py` 与 `covers.py` 后，`Word.Application` 成功打开 `report_final.docx` 并导出 140417 字节 PDF；契约测试新增 `test_generated_word_fields_are_word_compatible`，直接检查 XML 父子关系与 `pBdr` 顺序。

### 11.7 进程清理

每次 COM 调用都 `DispatchEx` 独立实例，关闭文档并 `Quit`。转换后立即检查可能看到 Word 进程收尾；等待数秒后 `tasklist` 无 `WINWORD/EXCEL/POWERPNT`。测试在 20 秒窗口内轮询，断言无新增残留进程。

## 遗留限制（本轮实测）

- dsoffice 对 Excel 渲染仍要求 `--sheet` + `--cell-range`，且拒绝 `--pages`；COM 回退用 print area 近似区域，分页/缩放可能与 dsoffice 不同。
- COM 回退需要 Windows + pywin32 + Office/WPS；Word 未暴露 `Hwnd/ProcessID`，退出可能滞后数秒，测试用 20 秒窗口确认。
- `validate_word_document` 的字体检查只看正文/页眉页脚显式 `w:rFonts`；只继承样式/主题、未显式写字体的 run 不解析最终字体。
- 标题编号只解析手工文字前缀（阿拉伯数字、中文数字、第 X 章/节）；Word 自动编号（`w:numPr`）不参与连续性判断。
- 中英文字体混用的判定是确定性规则：同段出现多个中文字体或多个西文字体，或纯中文 run 只设西文字体 / 纯英文 run 只设中文字体；不做字形覆盖分析。
- TOC 是否“需更新”看 `w:dirty` 或 `w:updateFields`；任一脏域可能让检查通过，未做字段级精确关联。
- `--reflow` 只处理幻灯片顶层形状的几何；不进入组合内部、不改 z-order/动画/旋转，也可能移动装饰形状；无法安全放置时保留原状并 warning。
- `--fix-spacing` 仍是旧的纵向间距逻辑，不处理越界/重叠；需要几何修复用 `--reflow`。
- PDF 转换是 dsoffice/LibreOffice 或 Office/WPS 的近似还原；COM 与 dsoffice 输出字节不同，不承诺与打印逐像素一致。
- 扫描件 OCR 依赖本机 Tesseract 与对应语言包；未安装时降级返回文本层。
- 活动文档入口只连接已打开的应用，不启动 Word/WPS；COM 转换回退是受控例外。
