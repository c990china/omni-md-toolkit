# NOTICE · 第三方组件与许可

本工具包依赖以下开源组件。各组件版权归其作者所有，并遵循其上游许可。
**MIT 许可证组件**在再分发时须保留其版权声明与许可声明。

## 核心引擎

| 组件 | 版权/来源 | 许可 | 用途 |
|---|---|---|---|
| [markitdown](https://github.com/microsoft/markitdown) | © Microsoft | MIT License | 通用格式转换引擎（xlsx/xls/csv/html/msg/rtf/odt/txt/md） |

markitdown 是本项目通用文本/数据格式转换的基础引擎；本项目的以下部分为在
markitdown 之上的自研增量代码（同样以 MIT 发布）：Office 合并单元格还原、
文本型 PDF 原生分栏抽取、Kindle 电子书解包抽取、中文版式自整理、
广告水印清理、34 条质量门禁、扫描件两段式直识流程。

## 其他依赖

| 组件 | 许可 | 用途 |
|---|---|---|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | AGPL-3.0 / 商业双许可 | 扫描件页面渲染、PDF 处理 |
| [pdfplumber](https://github.com/jsvine/pdfplumber) | MIT License | 文本型 PDF 解析 |
| [python-docx](https://github.com/python-openxml/python-docx) | MIT License | Word 文件解析 |
| [python-pptx](https://github.com/scanny/python-pptx) | MIT License | PowerPoint 文件解析 |
| [pandas](https://pandas.pydata.org/) | BSD-3-Clause | 表格降级与数据输出 |
| [openpyxl](https://openpyxl.readthedocs.io/) | MIT License | xlsx 读写 |
| [tabulate](https://github.com/astanin/python-tabulate) | MIT License | 表格渲染 |
| [requests](https://github.com/psf/requests) | Apache-2.0 | HTML 内联远程图片抓取 |
| [Pillow](https://python-pillow.org/) | HPND (MIT-CMU) | 图片处理 |
| [beautifulsoup4](https://www.crummy.com/software/BeautifulSoup/) | MIT License | XHTML DOM 渲染 |
| [lxml](https://lxml.de/) | BSD-3-Clause | XHTML 解析 |
| mobi | 见 PyPI 项目页 | Kindle 电子书解包 |
| pywin32（仅 Windows） | PSF-2.0 | 旧版 .doc 的 Word COM 转换 |

## 说明

- 上述许可信息以各组件官方仓库 / PyPI 页面为准，如有出入以官方为准。
- PyMuPDF 采用 AGPL-3.0 与商业双许可：在符合 AGPL-3.0 的开源项目中使用没有问题；
  如在闭源商业产品中分发，请向 Artifex 获取商业授权或改用其他渲染库。
