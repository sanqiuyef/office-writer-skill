# 字体、字号、页码与页面设置

来源：mcp-office-writer 0.6.0 的 `utils.py`、`word_writer.py`、`scenes.py`、`covers.py`、`pdf_writer.py`（2026-11）；中文字号映射沿用上游 `CN_FONT_SIZE_MAP`。

适用边界：中文办公文档（Word/PDF）的常用排版参数；不覆盖复杂域代码、邮件合并、双栏学术模板的全部细节。

## 中文字号 → pt

| 字号 | pt | 字号 | pt |
|:---|---:|:---|---:|
| 初号 | 42 | 小初 | 36 |
| 一号 | 26 | 小一 | 24 |
| 二号 | 22 | 小二 | 18 |
| 三号 | 16 | 小三 | 15 |
| 四号 | 14 | 小四 | 12 |
| 五号 | 10.5 | 小五 | 9 |
| 六号 | 7.5 | 小六 | 6.5 |

`FontConfig.size`、`ParagraphStyle.font_size`、`PdfFontConfig.cn_font_size` 都接受“小四”这类名称或直接数字。数字按 pt 处理。

## Word 场景排版

`create_word_document.py --scene <name>` 会套用 `scenes.py` 的预设；显式 `page_setup` / `header_footer` / `default_font` 覆盖场景默认值。

| scene | 用途 | 正文 | 标题 | 特点 |
|:---|:---|:---|:---|:---|
| `academic` | 学术论文 | 小四 | 三号/小三/四号 | 1.3 倍行距，首行缩进 |
| `report` | 报告/提案 | 小四 | 三号 | 自动封面（`show_cover`） |
| `official` | 公文 | 三号仿宋 | 二号小标宋 | GB/T 9704 风格 |
| `contract` | 合同 | 小四 | 三号 | 条款段落 |
| `resume` | 简历 | 小四 | 二号 | 紧凑页面 |
| `exam` | 试卷 | 五号 | 三号 | 题距 |
| `journal` | 期刊论文 | 五号 | 小四 | 测绘学报规范 |

常用参数：

```json
{
  "page_setup": {"paper_size": "A4", "orientation": "portrait",
                 "margin_top": 1.0, "margin_bottom": 1.0,
                 "margin_left": 1.18, "margin_right": 1.18},
  "header_footer": {"header_text": "季度报告", "footer_text": "第 {page} 页"},
  "default_font": {"name": "宋体", "size": "小四"}
}
```

边距单位是**英寸**（1 英寸 ≈ 2.54 cm）；PDF 的 `PdfPageSetup` 边距单位是**毫米**。

## 页眉页脚与页码

- `HeaderFooter.footer_text` 支持 `{page}` 占位；上游写入的是真正的 PAGE 域（`w:fldChar`/`w:instrText` 包在 `w:r` 中），Word 打开后显示实际页码。
- 需要“第 X 页 / 共 Y 页”时把 `{page}` 放在文本中，例如 `第 {page} 页`；需要总页数域时用 Word 手工插入 `NUMPAGES`。
- 目录：`SectionItem(type="toc", text="目录", toc_levels="1-3")` 插入 TOC 域并标记 `w:dirty="true"`；Word 打开后更新域。`validate_word_document` 会检查 TOC 是否存在且标记为需更新。
- 封面：`cover_config` 传 `title/subtitle/author/date/organization/extra_lines/recipe_id`；`scene=report` 默认生成封面。

## PDF 字体

`create_pdf_document.py` 的字体探测顺序（`pdf_writer.py`）：

1. `font_config.cn_font_path` 显式指定的字体文件；
2. `font_config.cn_font` 指定族名（`simsun`/`simhei`/`msyh` 等）；
3. `%WINDIR%\Fonts` 下的候选字体；
4. 仍找不到则遍历 Fonts 目录中任意 `.ttf/.ttc`。

`cn_font` 留空自动检测。`en_font` 支持 `Helvetica`/`Courier`/`Times`，其余字体走 fpdf2 的字体加载。字体文件缺失时 PDF 生成会失败，CLI 退出码 3。

## 版式检查

- `validate_word_document.py` 检查空白页、行距一致性、表格完整性、标题层级/编号、图片、分节、页眉页脚、TODO/TBD 标记，以及字体是否本机安装、PAGE/TOC 域、表格列数/合并结构、空段落比例、中英文字体混用；`errors>0` 时退出码 1。
- 新检查全部是 warning，不改变退出码；字体检查读取正文/页眉页脚显式 `w:rFonts`，用 Windows 字体注册表与字体目录比对（含宋体/SimSun 等中英别名）。
- 它仍不做逐像素视觉验收；字体替换后的字形、分页和页脚位置需要渲染或人工打开确认（见 `render_qa.md`）。
