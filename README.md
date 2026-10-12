# omni-md-toolkit · 全能通用 Markdown 转换工具包

**把任意文档一键转成干净、合规、可人工审阅的 Markdown。**

Word / Excel / PPT / PDF（文本型 + 扫描型）/ 图片 / 网页 / 邮件 / 旧版 `.doc` / **Kindle 电子书（azw3 / mobi / kf8 / epub）**——丢给它，产出排版工整、表格无损、图片规范落盘的同名 `.md`。

> 本项目的通用格式转换引擎基于微软开源的 [markitdown](https://github.com/microsoft/markitdown)（MIT），
> 并针对中文场景重写了大量本地增强逻辑（见下方「致谢」）。

---

## 它解决什么问题

市面上的转换工具各有各的坑，本工具包一次性把这些坑都踩完并修好了：

| 痛点 | 本工具包的做法 |
|---|---|
| Word 合并单元格被展平错位 | `gridSpan` / `vMerge` 跨行跨列**精准还原**，不是拍平成网格 |
| 旧版 `.doc` 没人管 | Windows + Word COM 自动转 `.docx` 再走修复链路 |
| PPT 转完只剩正文、**备注丢了** | 演讲者备注**同页抽取**，以引用块输出（备注即内容，缺备注＝丢稿） |
| 中文排版稀烂 | 内置 34 条规则的版式自整理 + 无损质量门禁 |
| 扫描件 PDF 要配 OCR 密钥 | **不做 OCR、不调网关**：脚本渲染页面，由当前大模型自己看图写稿 |
| markitdown 的 epub 分支会静默丢正文 | Kindle 电子书走**本地原生解包引擎**（实测修复「16 章正文全丢、仅剩文件清单」） |

---

## 支持格式

| 类别 | 格式 | 引擎 |
|---|---|---|
| 文档 | `.docx` | 本地增强引擎（合并单元格精准还原 + 图片落盘） |
| 文档 | `.doc`（Word 97-2003） | Word COM 转 docx → 走 docx 链路 |
| 演示 | `.pptx` | 本地引擎（形状排序 + 演讲者备注同页抽取 + 图片落盘） |
| 表格/数据 | `.xlsx` `.xls` `.csv` | markitdown 引擎（pandas 优雅降级） |
| 网页/邮件/其他 | `.html` `.msg` `.rtf` `.odt` `.txt` `.md` | markitdown 引擎 / 本地直读 |
| PDF | 文本型 `.pdf` | 本地原生分栏聚类抽取（多栏、编号标题定级、CID 清洗） |
| PDF | 扫描型 `.pdf` / 纯图片 | 渲染页面 → 当前模型看图写稿（**零 OCR、零密钥**） |
| 电子书 | `.azw3` `.mobi` `.kf8` `.epub` | 本地原生解包 + XHTML DOM 抽取 |

---

## 快速开始

```bash
# 1) 安装依赖（Python 3.8+）
pip install -r requirements.txt

# 2) 单文件转换：产出与源文件同名的 <基名>.md
python scripts/convert.py "工作报告.docx"

# 3) 批量转换整个文件夹（自动整理 + 门禁体检）
python scripts/pipeline.py "D:\待处理文档"
python scripts/pipeline.py "D:\待处理文档" --skip-exists   # 断点续跑
```

转换链路全自动：**内容转换 → 图片反解落盘 → 版式自整理 → 广告水印清理 → 34 条质量门禁**。
任何一步失败都会明确报错，**绝不把没整理的半成品当成品交付**。

### 扫描件 / 图片：两步走

```bash
python scripts/convert.py "扫描件.pdf"     # 第 1 步：渲染页面 + 打印作业单
# 按照作业单，由当前大模型逐页看图，把内容写进 _pages/<基名>/page_NN_ocr.md
python scripts/convert.py "扫描件.pdf"     # 第 2 步：重跑即自动合并 + 整理 + 门禁
```

支持断点续跑：已存在的逐页稿自动跳过，中断后随时继续。

---

## 核心设计：两段式交付

**第一段（脚本，确定、免费、可复现）** 管版式：合并断行、定标题层级、去页眉页脚、压空行。
**第二段（大模型，需要判断）** 管语义：拿原文对照，补回真丢的内容、按原文重建表格、修假标题。

```bash
# 第一段完成后，打印「四类待核对项」作业单（疑似缺失 / 表格问题 / 标题骨架 / 图片）
python scripts/llm_review.py "源文件.docx"
```

顺序不能反：模型在乱稿上改完，随后那次整理会把刚改对的东西又推翻一遍。

---

## 实用小工具

```bash
python scripts/md_reorganizer.py "a.md"        # 版式自整理（幂等，全链路只跑一次）
python scripts/md_lint.py "a.md" --source "源.docx"   # 34 条规则门禁：7 E + 23 W + 4 I
python scripts/fix_tables.py "a.md"            # 已转好的 MD 表格坏了？就地修复（默认预览，--apply 落盘）
python scripts/docx_pptx_to_md.py pptx "slides.pptx"   # PPT 逐页 + 演讲者备注同页抽取
python scripts/media_extract.py image "图目录" --alt-file "alt.txt"   # 图片批量抽离 + alt 作业单
python scripts/ad_clean.py "a.md" --ad-image   # 广告水印清理（图片广告由当前模型看图判断）
```

`标题黑名单.txt` 是**用户可维护的例外清单**：整理器靠启发式猜标题，总有几句正文短句被误升成标题——把它们登记进去，下次重跑就不再误判（比每次手工改产物更彻底）。

---

## 目录结构

```
omni-md-toolkit/
├── README.md                本文
├── LICENSE                  MIT
├── NOTICE.md                第三方组件与许可
├── requirements.txt
├── 标题黑名单.txt             # 标题误判的可持续压制入口（每行一条）
├── docs/
│   └── 版式自整理规范.md      # 排版目标态 + 34 条门禁规则表
└── scripts/
    ├── convert.py            # 单文件统一入口（路由派发）
    ├── pipeline.py           # 目录级批量一键流水线
    ├── md_reorganizer.py     # 版式自整理内核
    ├── md_lint.py            # 34 条质量门禁
    ├── docx_merge_table.py   # Word 合并单元格精准还原
    ├── docx_pptx_to_md.py    # PPT 逐页 + 演讲者备注抽取
    ├── doc_to_docx.py        # 旧版 .doc 经 Word COM 转 docx
    ├── extract_text_pdf.py   # 文本型 PDF 原生分栏抽取
    ├── kindle_to_md.py       # Kindle 电子书原生解包抽取
    ├── ocr_llm.py            # 扫描件两段式直识（渲染 + 作业单）
    ├── ad_clean.py           # 广告水印清理
    ├── media_extract.py      # 跨格式图片抽离落盘
    ├── fix_tables.py         # MD 表格事后修复
    └── llm_review.py         # 语义核对作业单（第二段入口）
```

---

## 实测记录

以下均为本工具包在真实交付中验证过的结果，不是推断：

- **GB/T 1.1—2020 标准 PDF**：多栏分栏聚类 + 编号标题定级直接抽取，目次区字体子集损坏时按真实章节白名单重建。
- **21 页 PPTX**：21 条演讲者备注（约 4000 字）全部同页抽取，正文 + 备注零丢失。
- **Kindle 电子书（azw3）**：`《目的与利润(精校本)》` 转换后门禁 **E0 / W0 / I0**，含字符多重集零丢失；章节顺序按 `content.opf` 的 spine，不按文件名字典序。
- **19 份 HR 流程图文档**：mermaid 图全部收敛为「A4 竖排单图」，最大 17 节点 / 最深 12 层。
- **一次完整交付基线**：含合并单元格的 Word 报表、21 页 PPTX（含备注）、扫描件电子书全链路实测见上。

踩过的坑（含代码级修法）沉淀在 `docs/` 与源码注释里，改代码前建议先读注释里的「实测」记录。

---

## 致谢

本项目站在这些开源工作的肩膀上：

- **[markitdown](https://github.com/microsoft/markitdown)（Microsoft，MIT License）**——通用格式转换引擎，负责 xlsx/xls/csv/html/msg/rtf/odt/txt/md 等格式的稳定转换（xlsx/csv 在 markitdown 缺失时由 pandas 降级）；docx/pptx/旧版 doc/PDF/电子书走本项目的本地引擎，均是在其思路之上的自研增量。
- [PyMuPDF](https://pymupdf.readthedocs.io/)（Artifex，AGPL-3.0 / 商业双许可）——扫描件页面渲染与 PDF 处理。
- [pdfplumber](https://github.com/jsvine/pdfplumber)（MIT）——文本型 PDF 解析。
- [python-docx](https://github.com/python-openxml/python-docx)（MIT）/ [python-pptx](https://github.com/scanny/python-pptx)（MIT）——Office 文件解析。
- [pandas](https://pandas.pydata.org/)（BSD-3-Clause）/ [openpyxl](https://openpyxl.readthedocs.io/)（MIT）/ [tabulate](https://github.com/astanin/python-tabulate)（MIT）——表格降级与数据输出。
- [requests](https://github.com/psf/requests)（Apache-2.0）/ [Pillow](https://python-pillow.org/)（HPND）——图片抓取与处理。
- [beautifulsoup4](https://www.crummy.com/software/BeautifulSoup/)（MIT）/ [lxml](https://lxml.de/)（BSD-3-Clause）——XHTML DOM 渲染。
- `mobi`（许可见 PyPI 项目页）——Kindle 电子书解包。
- `pywin32`（PSF-2.0，仅 Windows）——旧版 `.doc` 的 Word COM 转换。

完整许可清单见 [NOTICE.md](NOTICE.md)。

---

## 开源许可

本项目以 [MIT License](LICENSE) 发布，Copyright (c) 2026 c990china。

第三方组件各自遵循其上游许可；其中 PyMuPDF 为 AGPL-3.0 与商业双许可，
如在闭源商业产品中使用，请自行评估其商业授权要求。
