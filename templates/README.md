# templates/

三个可编辑的起步模板，由本 skill 生成，不含第三方正文。

| 文件 | 用法 |
|:---|:---|
| `word_template.docx` | `create_word_document.py --template templates/word_template.docx`；含页眉/页脚与 `{{title}}`、`{{summary}}`、`{{body}}`、`{{company}}`、`{{footer}}` 占位符 |
| `excel_template.xlsx` | `create_excel_workbook.py --template templates/excel_template.xlsx`；含“概览/明细”工作表和 `{{title}}`、`{{owner}}` 等占位符 |
| `powerpoint_template.pptx` | `create_powerpoint.py --template templates/powerpoint_template.pptx`；标题页 + 内容页，含 `{{title}}`、`{{subtitle}}`、`{{section}}`、`{{point1}}`、`{{point2}}` |

模板合并示例：

```bash
python -X utf8 scripts/merge_office_template.py \
  --template-path templates/word_template.docx \
  --output-path 输出.docx \
  --params '{"data":{"company":"示例公司","footer":"内部资料","title":"项目总结","summary":"摘要","body":"正文"}}'
```

注意：

- `create_*` 的 `--template` 是“继承样式/页面设置”，不是填充占位符；填充占位符用 `merge_office_template.py`。
- 占位符只替换 `{{key}}` 双花括号；嵌套对象用 `{{user.name}}`，数组索引用 `{{items[0].title}}`。
- 模板里的字体/配色可按项目要求自行修改；修改后建议跑一次 `validate_word_document.py` 和 `render_document_pages.py`。
