# assets/

可分发示例与运行时缓存目录。

## examples/

`examples/*.json` 是可直接传给 `--params-file` 的合成示例：

| 文件 | 用途 |
|:---|:---|
| `word_report.json` | 创建带标题、段落、表格的 Word 报告 |
| `excel_table.json` | 创建带冻结首行的 Excel 工作表 |
| `ppt_deck.json` | 创建标题页 + 内容页 PPT |
| `framework_diagram.json` | 生成三阶段框架图 |

示例中的 `output_path` 是相对路径，运行前按需改成目标目录。示例文本全部为自制合成内容，不含第三方论文或客户资料。

## image_cache/

`create_powerpoint` 插入 `url` 图片时，上游会把图片缓存到 `assets/image_cache/`（同 URL 复用，单张 20MB 上限）。该目录运行时创建，已在 `.gitignore` 中排除，不随 skill 分发。

## 分发边界

- 不要把客户文档、未公开数据、第三方全文放进 `assets/`；
- 需要项目专用模板时，放 `templates/` 或项目自己的目录，并在交付说明中写明来源与授权；
- 本 skill 的模板和示例是通用起步文件，不承诺适配特定学校/期刊/公司的格式规范。
