# templates/

全部模板由本 skill 生成，只含自制合成文案，不含第三方正文或客户数据。可分发素材的统一索引见 `assets/manifest.json`。

## 基础模板（骨架 + 占位符）

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

## PPT 模板库（8 套 16:9 成品模板）

每套两个文件：

- `<slug>.pptx`：8 页示例 deck，直接打开改文字即可交付；
- `<slug>.potx`：0 页版式模板，供 `create_powerpoint --template` 生成新 deck，不会带出示例页。

| slug | 名称 | 场景 | 适用 | 字体（标题/正文） |
|:---|:---|:---|:---|:---|
| `business-report` | 商务汇报 | 商务办公 | 年度/季度经营分析、项目复盘、部门总结 | 微软雅黑 / 微软雅黑 |
| `consulting` | 咨询提案 | 商务办公 | 战略咨询、方案提案、投融资路演 | 微软雅黑 / 微软雅黑 |
| `academic-defense` | 学术答辩 | 学术科研 | 学位论文答辩、课题结题、会议报告 | 华文中宋 / 微软雅黑 |
| `academic-seminar` | 开题与组会 | 学术科研 | 开题报告、组会、中期检查 | 华文中宋 / 微软雅黑 |
| `government-report` | 政务汇报 | 政务公文 | 工作汇报、年度总结、督察整改 | 华文中宋 / 仿宋 |
| `government-briefing` | 政策宣讲 | 政务公文 | 政策解读、宣讲培训、落实部署 | 华文中宋 / 仿宋 |
| `hbut-defense` | 湖工大学术答辩 | 学术科研 | 湖工大学位答辩、开题、结题 | 微软雅黑 / 微软雅黑 |
| `competition-defense` | 竞赛答辩 | 竞赛路演 | 挑战杯、互联网+、项目答辩 | 华文中宋 / 微软雅黑 |

统一页序：封面 → 目录 → 章节页 → 内容页 → 图表页 → 表格页 → 两栏页 → 结尾。配色、页面用途与每页示例文案见 `assets/manifest.json`。

### 用法一：套版式生成新 deck（推荐）

```bash
python -X utf8 scripts/create_powerpoint.py \
  --output-path 答辩.pptx \
  --template templates/ppt/academic-defense.potx \
  --params '{"slides":[
    {"layout":"title","title":"论文题目","subtitle":"答辩人 · 导师 · 日期"},
    {"layout":"content","title":"研究背景","body":[{"type":"paragraph","text":"..."}]}
  ]}'
```

`.potx` 是 0 页版式模板，生成的 deck 只包含 `slides` 里定义的内容；标题/正文继承版式的字号、颜色与对齐（标题 30pt 主题色左对齐，正文 18pt）。

### 用法二：复制示例 deck 改文字

```bash
cp templates/ppt/business-report.pptx ./2026年度汇报.pptx
# 再用 edit_powerpoint 追加、batch_edit_office_document 改文字，或直接在 PowerPoint 中编辑
```

### 用法三：填充 `{{占位符}}`

`merge_office_template.py` 支持 `.docx` / `.xlsx` / `.pptx`，只替换 `{{key}}` 文本；示例 deck 里的文案是真实示例而非占位符，需要替换时按文本查找改。

## 重新生成 PPT 模板

```bash
python -X utf8 tools/build_ppt_templates.py
```

脚本会重写 `templates/ppt/` 下全部 16 个文件（8 套 × pptx/potx），并保持 11 个标准版式的顺序——`create_powerpoint` 按固定索引映射版式，顺序变化会导致内容落到错误的版式。改配色、字体或示例文案都改这个脚本，不要手工编辑二进制模板。

## 约束

- 8 套 PPT 模板均为 16:9（13.333 × 7.5 英寸），与 `create_powerpoint` 的默认画幅（10 × 7.5）不同；用 `.potx` 新建时画幅随模板走。
- 标题字体依赖华文中宋（`STZHONGS.TTF`）或微软雅黑；政务两套正文用仿宋。目标机器缺字体时 PowerPoint 会回退，版式不保证一致。
- `hbut-defense` 与 `competition-defense` 只含校色、版式与示例文案；校徽、校园实景图、赛事标识均为虚线占位框，学校版权素材不分发，使用前自行放入。
- `create_*` 的 `--template` 是“继承样式/版式”，不是填充占位符；填充占位符用 `merge_office_template.py`。
- 占位符只替换 `{{key}}` 双花括号；嵌套对象用 `{{user.name}}`，数组索引用 `{{items[0].title}}`。
- 模板里的字体/配色可按项目要求自行修改；修改后建议跑一次 `validate_word_document.py`（Word）或 `convert_to_pdf.py` + `render_document_pages.py`（PPT/Excel）复核。
