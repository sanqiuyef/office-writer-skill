# WPS / COM 活动文档

来源：mcp-office-writer 0.6.0 的 `office_com.py`（2026-11）、本 skill 的 `office_convert.py` COM 回退（2026-11-15 实测）与 `COM活动文档写入说明.md`。

适用边界：Windows + pywin32。活动文档入口只操作**已经打开**的 Word/WPS、Excel/WPS 表格、PowerPoint 窗口；转换/渲染/重算的 COM 回退会启动独立隐藏实例并退出。不负责修复损坏的 COM 注册表。

## 两种 COM 用法

| 场景 | 入口 | 是否启动应用 | 目标 |
|:---|:---|:---|:---|
| 活动文档编辑 | `get_open_office_documents`、`edit_open_*`、`ppt_add_flowchart_slide` | 否 | 用户已打开的窗口 |
| 文件转换/渲染/重算 | `convert_to_pdf`、`render_document_pages`、`recalculate_excel_formulas` | 是（`DispatchEx` 隐藏实例） | 磁盘文件 |

转换回退按扩展名选 ProgID：Word（`Word.Application` → `KWps.Application` → `Word.Application`）、Excel（`Excel.Application` → `KET.Application` → `Excel.Application`）、PowerPoint（`PowerPoint.Application` → `KWPP.Application` → `PowerPoint.Application`）。每次调用都关闭文档并 `Quit`；Word 进程可能滞后数秒退出，拿到 PID 时会等待并在必要时 `taskkill /F /T`。活动文档入口不会碰这些隐藏实例。

## 活动文档能力

| 入口 | 行为 | 不启动应用 |
|:---|:---|:---|
| `get_open_office_documents.py` | 列出 Word.Application / kwps.Application、Excel.Application / et.Application、PowerPoint.Application 中已打开的文档 | 是 |
| `edit_open_word_document.py` | 向活动 Word/WPS 文档追加标题、段落、列表、分页符 | 是 |
| `edit_open_excel_workbook.py` | 向活动 Excel/WPS 表格追加工作表 | 是 |
| `edit_open_powerpoint.py` | 向活动 PowerPoint 演示文稿追加幻灯片 | 是 |
| `ppt_add_flowchart_slide.py` | 在活动 PowerPoint 中追加一张流程图幻灯片 | 是 |

`edit_word_document.py`、`edit_excel_workbook.py`、`edit_powerpoint.py` 会先尝试活动 COM 文档；找不到活动文档才走磁盘文件。因此同样的命令在“文件已打开”和“文件未打开”时可能走不同路径，结果消息也不同。

## 超时与回退

所有活动文档入口都有 60 秒上限（`config/defaults.json` 的 `timeouts.com_operation_seconds`）：

- 超时不会强行结束 Office 进程，而是返回成功为 false 的 `EditResult`，CLI 退出码 3；
- `get_open_office_documents.py` 超时返回 `{"word": [], "excel": [], "powerpoint": []}`，仍是成功；
- 四个编辑入口在调用前先 `com_is_available()`；不可用时返回明确错误，不抛裸异常；
- 找不到目标文件对应窗口时返回“目标文件未在活动 ... 实例中打开”。
- 转换/渲染/重算的 COM 回退使用 300 秒上限（`timeouts.dsoffice_operation_seconds`）；超时由 `with_timeout` 返回失败结果，隐藏实例可能仍在收尾，随后检查 `tasklist`。

## 常见坑

1. **同名文件**：COM 按规范化绝对路径匹配；相对路径、短路径、大小写差异都可能导致匹配失败。CLI 会把参数原样交给上游，建议传绝对路径。
2. **只读文档**：活动文档为只读时上游直接拒绝保存；先关闭只读或另存副本。
3. **WPS 与 Office 同时运行**：两者可能都注册同一个 ProgID。活动文档的 `application` 字段显示实际命中的 ProgID；转换回退的 `com_engine` 按应用 `Path` 判定（含 `WPS` 即 WPS），不要只看 ProgID。
4. **弹窗/未响应**：保存确认、宏提示、字体替换提示都会让调用挂起。活动文档入口 60 秒超时；转换回退 300 秒超时。先处理弹窗或改用 dsoffice。
5. **回滚**：Excel/PPT 编辑失败时上游会删除本次新增的工作表/幻灯片；Word 使用 UndoRecord 回滚。若回滚也失败，错误消息会提示“请勿保存该文档”。
6. **Explorer 0 字节占位文件**：新建后尚未首次保存的 docx/xlsx 可能是 0 字节；上游会先尝试活动文档，再走 OOXML 校验。不要用 `read_*` 读取这种占位文件。
7. **无 Office 环境**：`com_is_available()` 为 False；活动文档与 COM 转换都不可用，但 dsoffice 仍可完成文件级转换/渲染/重算。
8. **Word 对 OOXML 更严格**：`w:fldChar`/`w:instrText` 必须包在 `w:r` 中，`w:pBdr` 必须位于 `w:spacing`/`w:ind`/`w:jc` 之前，否则 Word 报“打开文件时遇到错误”而 WPS/LibreOffice 仍能打开。本 skill 已在 `word_writer.py`/`covers.py` 修正这两类顺序。
9. **进程清理**：转换回退用 `DispatchEx` 独立实例；成功/失败都会 `Close` + `Quit`。Word 没有暴露 `Hwnd/ProcessID`，进程退出可能滞后数秒；测试等待 20 秒后断言无 `WINWORD/EXCEL/POWERPNT` 残留。

## 最小排查顺序

1. `get_open_office_documents.py` 是否能看到目标文件？
2. `application` 是 Office 还是 WPS？
3. 文档是否只读？是否有未处理弹窗？
4. 换绝对路径重试；
5. 仍失败则改用文件级 `edit_*`，并确认没有其他进程占用文件。
