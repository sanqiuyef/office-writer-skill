# Office Writer

把 mcp-office-writer 0.6.0 的 32 个文档工具拆成 32 个独立 CLI 入口，另加 12 个本地入口：5 个 Python 入口（活动文档替换、Word 批注/属性、Excel 图表、PPT 超链接）与 7 个 office-native 原生入口（Word 修订/脚注/内容控件、PPT 动画切换批注/SmartArt、Excel 透视表/打印），共 44 个入口。44 个入口默认由 C# `tools/office-native` 执行（Open XML SDK + Office COM）；`OFFICE_NATIVE_ENGINE=python` 可回退到 `scripts/office_engine/` 的 vendored 引擎做对照或排障。每个入口只做参数解析、契约校验和 JSON 输出。C# 后端构建：`pwsh -File tools/office-native/build.ps1`。skill 不依赖 MCP，不启动 MCP 服务。

Word / Excel / PPT 的逐项功能面与覆盖状态见 `references/capabilities_word.md`、`capabilities_excel.md`、`capabilities_ppt.md`（2026-11 扩展）。

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

## 写入安全

- 所有写文件的入口默认 `--if-exists error`：目标已存在且非空时拒绝写入，退出码 2，不触碰原文件。
- 空白文件（不存在或大小为 0）可直接写入，用于新建和占位符场景。
- 确需覆盖时同时传 `--if-exists overwrite --confirm-overwrite`；需要保留旧文件时用 `--if-exists rename`。
- 原地写回（`edit_*`、`batch_edit_office_document`、PPT 配色/布局/清理）在替换前自动生成 `*.bak-YYYYmmdd-HHMMSS` 备份，路径在 `data.backup_path`（或 `backup_path` 字段）返回。
- PDF 生成、PDF 转换、页面渲染覆盖已有目标时同样先备份；目录目标整体改名为 `.bak-*`，不删除旧内容。
- 已打开的文档优先走 COM 入口，由应用自己保存，保留公式、修订、批注、域和嵌入对象。

## 隐私

- 文档只在本机读写；不发送到任何 API，不上传遥测。
- 唯一的网络访问是 `create_powerpoint` 插入 `url` 图片时的下载（仅 http/https，20MB 上限，缓存到 `assets/image_cache/`）。
- dsoffice、Office/WPS COM、Tesseract 都是本地调用，不外发文档内容。
- 不要把客户文档、未公开数据、密钥写进 `templates/`、`assets/` 或提交记录。

## 测试

默认走 C# 后端：

```bash
python -X utf8 -m unittest discover -s tests -v
```

回退 Python 引擎对照：

```bash
OFFICE_NATIVE_ENGINE=python python -X utf8 -m unittest discover -s tests -v
```

2026-11-20 实测：默认 C# 套件 **58 项通过、2 项跳过**；`OFFICE_NATIVE_ENGINE=python` 套件 **58 项全通过**。跳过项是 fake dsoffice 的 `.cmd` shim 在当前沙箱无法启动 `cmd.exe`（普通 Windows 终端可执行）。覆盖 44 个入口 `--help`、固定输出全量比对、退出码 1/2/3、创建→读取→校验往返、结构/触发基线、字体回退、渲染计数修正、批量备份、PPT 越界/重叠审计与 `--reflow`、dsoffice 与真实 Office/WPS COM 转换/重算、COM 缺失回退、17 项 validate 实检 fixture、扫描件 OCR、覆盖闸门，以及 4 项 P2 回归（Excel 日期/流式读取、批量合并区、PPT 图表内嵌工作簿、Word 页眉样式与分节）。有意改固定 CLI 行为后先运行：

```bash
python -X utf8 tests/update_expected.py
```

真实执行记录、转换/渲染结果与遗留问题见 `tests/README.md`。

历史验证记录（2026-11-15/19：COM 转换、validate 17 项、reflow、覆盖闸门与能力扩展）见 git 历史与 `tests/README.md`。

## 维护

- `scripts/office_engine/*.py` 是 mcp-office-writer 0.6.0（2026-11）的 vendored 副本；每个文件头写明上游路径、版本和修改点。升级上游时整包替换并同步 `tests/test_parity.py` 的 golden。
- 本 skill 对引擎的增量：转换/渲染/重算回退链与 COM、17 项 validate 实检、PPT `reflow_slide_layout`、Word 域 XML 与封面边框顺序修正；2026-11 扩展 Word 超链接/通用域/文档属性、Excel 图表/条件格式/数据验证/超级表/名称/批注/图片/页面设置、PPT 表格与图表几何/图片裁剪/真列表/双栏占位符/`.potx`；44 个入口已全量迁移到 C# `tools/office-native/`（Open XML SDK + COM，默认引擎），旧 Python 引擎保留为 `OFFICE_NATIVE_ENGINE=python` 对照（源码在 `tools/office-native/`，二进制构建后放 `dist/`，不提交）。
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
- 2026-11-20 已补齐：Excel 日期/时间读取与流式读取、合并区批量写入保护、图表/图片/批注锚点边界、活动 Excel 复用、模板公式与名称清理、工作表保护与多键真实排序；Word Header/Footer 样式、嵌套/行级修订、跨 run track-delete、原子写回、书签、分栏/页面边框/文字水印；PPT 图表内嵌数据工作簿、`a:br`/`a:fld` 文本语义、跨 slide 关系迁移、节/放映设置/页脚页码。
- 仍未封装：Excel 形状/迷你图/工作簿结构保护/加密/Power Query；Word 邮件合并/比较合并/打印/浮动图片；PPT 母版/自定义放映/媒体跨页移动/视频导出。

## 许可

MIT，Copyright (c) 2026 sanqiuyef。`scripts/office_engine/` 来自 mcp-office-writer 0.6.0（MIT）。第三方许可与运行时依赖提示（PyMuPDF 为 AGPL-3.0/商业双许可，fpdf2 为 LGPL-3.0）见 `THIRD_PARTY_NOTICES.md`。
