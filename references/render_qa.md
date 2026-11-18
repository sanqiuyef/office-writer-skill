# 渲染验收：页面预览与字号/字体检查

来源：mcp-office-writer 0.6.0 的 `office_convert.py`、`reader.py` 与本地 pymupdf span 检查实践（2026-11）；COM/PyMuPDF 回退为 2026-11-15 实测。

适用边界：需要“看起来对不对”的验收，而不是仅检查 XML 结构。渲染优先 dsoffice，缺失时用 Office/WPS COM 转临时 PDF 再交给 PyMuPDF；PDF span 检查依赖 PyMuPDF。两者都是只读检查，不修改源文件。

## 渲染页面

```bash
python -X utf8 scripts/render_document_pages.py \
  --input-path 报告.docx \
  --output-dir 报告_pages \
  --pages 1,2 \
  --dpi 144
```

- 支持 docx/xlsx/pptx、97-2003 二进制格式与 pdf。
- `pages` 为 `all` 或 `1,3,5`；Excel 可用 `--sheet` + `--cell-range A1:D20` 只渲染区域。
- `dpi` 默认 144，上游会夹到 72–600；高 DPI 更清晰但更慢。
- 输出目录默认是源文件同目录的 `<文件名>_pages`；`if_exists=overwrite` 会先清空重建，`error` 直接报错。
- 返回 `images[]`（每页路径与尺寸）、`page_count`、`missing_fonts[]`；`page_count` 按实际 PNG 产物修正，上游元数据不一致时另见 `page_count_reported`。
- dsoffice 缺失/失败时：PDF 直接由 PyMuPDF 渲染；Office 文档先经 COM 转临时 PDF，再按 `--pages`/`--dpi` 出图。返回 `backend=pymupdf` 或 `com+pymupdf`，并带 `fallback_from[]`。

## 用 pymupdf 读 span 字号

渲染图只能看整体；要量化字号/字体，先 `convert_to_pdf.py` 转 PDF，再用 PyMuPDF 读文本 span：

```python
import pymupdf
doc = pymupdf.open("报告.pdf")
for page in doc:
    data = page.get_text("dict")
    for block in data["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                print(round(span["size"], 2), span["font"], span["text"][:40])
```

检查点：

1. 正文字号是否一致（同名样式不应出现 10.5 / 12 混排）；
2. 标题是否比正文大且层级递减；
3. 中文字体是否被替换成方框/黑体替代；
4. 表格是否跨页断裂、表头是否重复；
5. 页脚页码是否存在且连续；
6. 图片是否被裁切、是否超出页边距。

`read_pdf_structured.py --include-blocks` 也能返回每页 bbox/字体/字号/颜色，适合脚本化抽查，但大文档响应体较大。

## 缺失字体

`convert_to_pdf.py` 返回 `missing_fonts[]`；dsoffice 会按字形度量替换，版式通常接近原样，但字形会变。验收时：

- 缺字体不影响“是否可读”，但影响“是否与源文档一致”；
- 对外交付前，在源机器安装字体或改用通用字体（宋体/黑体/微软雅黑/Arial）；
- 不要把 `missing_fonts` 当成转换失败；它只是提示。

## 推荐验收顺序

1. `validate_word_document.py` 结构门控（errors=0，并查看字体/标题编号/PAGE/TOC/表格/空段落/中英混排 warning）；
2. `convert_to_pdf.py` 转 PDF，确认 `missing_fonts`；
3. `render_document_pages.py` 抽首尾页 + 表格页；
4. 用 pymupdf span 检查字号/字体；
5. 人工看一遍版式（缩进、行距、分页、页眉页脚）。

## 边界

- 渲染是“视觉近似”，不保证与 Word/WPS 打印结果逐像素一致；
- PDF span 的 `size` 是 PDF 实际字号，可能与 Word 样式字号因缩放/替换略有差异；
- 扫描件 PDF 的文本层可能为空，`reader.py` 会在满足条件时尝试 OCR（需要 Tesseract；见 `conversion.md`）。
