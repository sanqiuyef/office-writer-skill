# Office Writer

把 mcp-office-writer 0.6.0 的 32 个文档工具拆成 32 个独立 CLI 入口。每个入口只做参数解析、Pydantic 校验和 JSON 输出，文档逻辑来自 `scripts/office_engine/` 的 vendored 引擎。skill 不依赖 MCP，不启动 MCP 服务。

## 安装

作为 ZCode skill 使用：把本目录放到 `~/.zcode/skills/office-writer`（用户级）或项目的 `.zcode/skills/office-writer`（项目级），重启 agent 后生效。其他支持 `SKILL.md` 的 agent 按目标平台的 skill 目录放置即可。

运行时使用任意 Python ≥ 3.10；首次使用先安装依赖：

```bash
pip install -r requirements.txt
```

`requirements.txt` 不含 mcp 依赖；`scripts/` 内也没有 MCP 导入。转换、渲染与公式重算至少启用一种后端：

1. **dsoffice**：`npm i -g dsoffice`（Node.js ≥ 22.19；Windows x64 另需 MSVC v14 运行库），用 `which dsoffice` 检测。
2. **Microsoft Office COM**：Windows 安装 Word/Excel/PowerPoint 与 pywin32，设 `OFFICE_CONVERT_BACKEND=com`（或 `auto`）。
3. **WPS Office COM**：Windows 安装 WPS（`KWps.Application`/`KET.Application`/`KWPP.Application`）与 pywin32，设 `OFFICE_CONVERT_BACKEND=com`（或 `auto`）。

`OFFICE_CONVERT_BACKEND=auto|dsoffice|com|none` 控制顺序：默认 `auto` 先 dsoffice，缺失或执行失败回退 Office COM，再回退 WPS COM；`none` 禁用并在错误中列出三种启用方式。扫描件 OCR 需要本机 Tesseract，可用 `OFFICE_TESSERACT_CMD` 指定可执行文件（兼容 `TESSERACT_CMD`）。

## 运行时

- Windows 10/11 + Python ≥ 3.10；COM 入口需要 pywin32 和本机 Office/WPS。
- Office 文件读写用 python-docx/openpyxl/python-pptx；PDF 读取/渲染用 PyMuPDF，PDF 生成用 fpdf2。
- 所有入口默认 UTF-8 stdio；`--help` 正常打印帮助，其他情况 stdout 只输出 JSON。
- `config/defaults.json` 可调超时和渲染默认 DPI；同名环境变量见 `.env.example`。

## 隐私

- 文档只在本机读写；不发送到任何 API，不上传遥测。
- 唯一的网络访问是 `create_powerpoint` 插入 `url` 图片时的下载（仅 http/https，20MB 上限，缓存到 `assets/image_cache/`）。
- dsoffice、Office/WPS COM、Tesseract 都是本地调用，不外发文档内容。
- 不要把客户文档、未公开数据、密钥写进 `templates/`、`assets/` 或提交记录。

## 测试

```bash
python -X utf8 -m unittest discover -s tests -v
```

本轮 **42 项通过、0 跳过**。覆盖 32 个入口 `--help`、固定输出全量比对、退出码 1/2/3、创建→读取→校验往返、结构/触发基线、字体回退、渲染计数修正、批量备份、PPT 越界/重叠审计与 `--reflow`、dsoffice 与真实 Office/WPS COM 转换/重算、COM 缺失回退、17 项 validate 实检 fixture、扫描件 OCR。有意改固定 CLI 行为后先运行：

```bash
python -X utf8 tests/update_expected.py
```

真实执行记录、转换/渲染结果与遗留问题见 `tests/README.md`。

## 验证记录（2026-11-15）

- 全量测试：42 项通过、0 跳过。32 个入口 `--help` 全部退出 0。
- dsoffice：docx/xlsx/pptx → PDF 均退出 0；docx 报告 `Courier` 缺失，回退建议 `Consolas`。
- COM：docx→`Word.Application`、xlsx→`Excel.Application`、pptx→`PowerPoint.Application` 均退出 0；`com_engine=office`。WPS `KWps/KET/KWPP` 也实测可导出。
- 渲染回退：`MODULE_OFFICE_DSOFFICE` 指向不存在路径时，docx 先经 COM 转 PDF 再由 PyMuPDF 出图；`--pages 1 --dpi 144` 生成 1191×1684 PNG，`page_count=1`。
- 公式重算：`OFFICE_CONVERT_BACKEND=com` 走 `CalculateFullRebuild`+`Save`，回读合计列 `255/390`。
- validate：干净 fixture 17/17 通过；问题 fixture 命中 8 类 warning（字体、标题层级/编号、PAGE、TOC、表格列、空段落、中英混排）；TOC 未标脏 fixture 单独命中。
- reflow：越界、重叠、混合、超大四类形状重排后 `remaining_warning_count=0`；例如重叠形状 B 从 `(3.0,2.0)` 移到 `(3.0,3.15)`，越界形状 C 从 `(9.0,6.0)` 移到 `(6.9,5.4)`。
- 每次 COM 调用后等待 Quit；测试断言无残留 `WINWORD/EXCEL/POWERPNT` 进程。

## 维护

- `scripts/office_engine/*.py` 是 mcp-office-writer 0.6.0（2026-11）的 vendored 副本；每个文件头写明上游路径、版本和修改点。升级上游时整包替换并同步 `tests/test_parity.py` 的 golden。
- 本 skill 对引擎的增量：转换/渲染/重算回退链与 COM、17 项 validate 实检、PPT `reflow_slide_layout`、Word 域 XML 与封面边框顺序修正。
- 入口脚本只保留 `build_parser()`/`run(args)`/`main()`；共享契约在 `scripts/_common.py`。新增能力应新增 1:1 入口，不把逻辑塞进 `_common.py`。
- 素材：`templates/` 起步模板、`assets/examples/` 合成示例、`references/` 主题文档。模板/示例改动后跑测试与一次渲染验收。
- 超时/DPI 默认值在 `config/defaults.json`；不要把密钥或机器绝对路径写进配置。

## 已知边界（已解决/降级）

- dsoffice 不再是硬依赖：转换/渲染/重算会回退 Office/WPS COM；两者都不可用时退出码 3，错误列出三种启用方式。
- COM 回退需要 Windows + pywin32 + Office/WPS；转换用独立 `DispatchEx` 实例并 Quit，不影响已打开文档。Word 进程退出可能滞后数秒，测试等待 20 秒。
- COM 回退的 Excel 渲染用 sheet + print area 近似 `--cell-range`；复杂分页/缩放仍可能与 dsoffice 不同。
- `validate_word_document` 已实检字体安装、标题层级/编号、PAGE/TOC 域、表格列/合并、空段落比例、中英文字体混用；仍不做逐像素视觉验收。
- `ppt_optimize_slide_layout` 默认只审计；`--reflow` 自动重排越界/重叠（0.1in 边距、0.15in 间距），`--fix-spacing` 仍是旧的纵向间距优化。无法安全放置的形状保留原状并给 warning。
- PDF 转换是 dsoffice/LibreOffice 或 Office/WPS 的近似还原，不承诺与打印逐像素一致。
- 扫描件 OCR 依赖本机 Tesseract；未安装时降级返回文本层并在 `message` 说明。
- 文件被 Office 占用、损坏 OOXML、加密文档等由底层库报错；CLI 统一转成 JSON 错误，不自动修复源文件。

## 许可

MIT，Copyright (c) 2026 sanqiuyef。`scripts/office_engine/` 来自 mcp-office-writer 0.6.0（MIT）。运行时依赖的许可提示见 `LICENSE`，其中 PyMuPDF 为 AGPL-3.0/商业双许可，fpdf2 为 LGPL-3.0。
