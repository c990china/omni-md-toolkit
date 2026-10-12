---
name: omni-md-toolkit
description: |
  全能通用全格式 Markdown 转换与自整理工具包（基于微软 markitdown + 本地自研增强引擎）。
  支持 Word / Excel / PPT / PDF（文本型+扫描型）/ 图片 / 网页 / 邮件 / 旧版 .doc / Kindle 电子书
  （azw3/mobi/kf8/epub）一键转成同名 .md：合并单元格精准还原、PPT 演讲者备注同页抽取、
  中文版式自整理、34 条无损质量门禁、扫描件由当前模型看图逐页写稿（零 OCR、零网关）。
  当用户提出「格式转 md」「把这个文件夹里的文件转 md」「整理 md」「根据：... 转 md」时使用。
version: 1.3.0
---

# omni-md-toolkit · 全能通用 Markdown 转换工具包

> 把任意文档一键转成干净、合规、可人工审阅的 Markdown。
> 详细说明见仓库根目录 `README.md`。

## 用法（交给 AI）

**单文件转换**（自动整理 + 门禁体检，只产出 `<基名>.md` 一份）：

```bash
python scripts/convert.py "文件.docx"        # 自动识别类型
python scripts/convert.py "扫描件.pdf"       # 渲染页面 + 打印作业单，重跑自动合并
python scripts/convert.py "电子书.azw3"      # 本地原生解包，不走 markitdown
```

**批量流水线**：

```bash
python scripts/pipeline.py "文件夹"                  # 标准：转换 + 整理 + 门禁
python scripts/pipeline.py "文件夹" --skip-exists    # 断点续跑
python scripts/pipeline.py "文件夹" --reorganize-only # 仅整理已有 MD
```

**两段式交付**：第一段脚本管版式（`md_reorganizer.py`），第二段模型管语义
（`llm_review.py` 打印「疑似缺失 / 表格问题 / 标题骨架 / 图片」四类待核对项）。
**顺序不能反**：模型改完之后不要再跑整理器，只用 `md_lint.py` 复检。

## 铁律

1. **产物只一份**：与源文件同名的 `<基名>.md`，不产出 `_整理` 副本；
   自整理失败时**宁可报错不产出**，不许裸稿冒充成品。
2. **全链路只整理一次**（整理器不幂等，连跑会周期振荡）。
3. **扫描件零 OCR、零网关**：脚本只渲染页面，识别由当前模型看图完成；
   判定「没有视觉能力」前先单张 Read 试一次。
4. **电子书不走 markitdown**：epub 分支缺 `ebooklib` 时会静默只输出 zip 清单、
   正文一字不抽（实测 16 章全丢），必须走 `kindle_to_md.py`。
5. **门禁不等于语义核对**：`md_lint.py` 只查格式与无损，「该有的标题/图片有没有」
   要靠 `llm_review.py` 第二段。
6. **标题黑名单**：`标题黑名单.txt`（仓库根目录或环境变量 `TITLE_BLACKLIST` 指定）
   是可持续的误判压制入口，登记一次永久生效。

## 文件地图

| 脚本 | 职责 |
|---|---|
| `convert.py` | 单文件统一入口（路由派发） |
| `pipeline.py` | 目录级批量一键流水线 |
| `md_reorganizer.py` | 版式自整理内核 |
| `md_lint.py` | 34 条规则质量门禁（7 E + 23 W + 4 I） |
| `docx_merge_table.py` | Word 合并单元格精准还原 |
| `docx_pptx_to_md.py` | PPT 逐页 + 演讲者备注同页抽取 |
| `doc_to_docx.py` | 旧版 .doc 经 Word COM 转 docx |
| `extract_text_pdf.py` | 文本型 PDF 原生分栏抽取 |
| `kindle_to_md.py` | Kindle 电子书原生解包抽取 |
| `ocr_llm.py` | 扫描件两段式直识 |
| `ad_clean.py` / `media_extract.py` / `fix_tables.py` | 广告清理 / 图片抽离 / 表格修复 |
| `llm_review.py` | 语义核对作业单（第二段入口） |

踩坑与实证记录见各脚本顶部注释与 `docs/`。
