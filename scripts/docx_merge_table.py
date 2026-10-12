# -*- coding: utf-8 -*-
"""DOCX -> Markdown 精确转换（修正合并单元格；docx_pptx_to_md.py 的替代方案）。

## 为什么需要这个脚本

`docx_pptx_to_md.py` 用 python-docx 的 `row.cells` 取单元格，而该 API 会把
Word 的合并单元格「展开」：
- `w:gridSpan`（横向合并 N 列）-> 同一个值重复 N 次（`姓名 | 代立信 | 代立信`）
- `w:vMerge`（纵向合并）-> 同一行重复出现多次
且它**不输出 `| --- |` 表头分隔行**，产出的并不是合法 Markdown 表格。
`md_reorganizer.py` 也修不了（其 detect_table 只识别 PDF 文本型表格）。
简历、人员信息表这类「通篇大表格」的文档受影响最严重。

## 本脚本做法

- 直接遍历 OOXML：`table._tbl.tr_lst` -> `tr.tc_lst`，绕开 `row.cells`
- `w:gridSpan`：一个 tc 占 N 列，渲染时**折叠相邻重复值**（而不是铺满 N 列）
- `w:vMerge`：`val="restart"` 为起始，其余为续接；续接单元格继承上方同列的值
  （继承后会出现完全相同的相邻行，再统一去重）
- 输出标准 Markdown 表格（含 `| --- |` 分隔行）
- 去尾部空列、去全空行、去「空壳行」（除首列外只剩 `2、` `3、` 这类孤立编号的碎片行，
  是 Word 里内容被拆散后的残留）
- 单元格内多段落用 `<br>` 连接；`|` 转义为 `\\|`
- 标题按 Word `Heading N` 样式映射 `#`；首行短句自动提升为 H1
- 图片：从 `a:blip` 的 `r:embed` 取 `part.related_parts` 的 blob，落盘到
  `assets_<基名>/` 并按原位置内联

## 用法

    python *_docx_merge_table.py "a.docx" [-o "a.md"] [--assets-dir DIR]

## 风险提示

「折叠相邻重复单元格」基于启发式：若同一行里两个**相邻且内容完全相同**的单元格
在原文中确有语义区别，会被合并。对简历/表单类文档安全；对统计表请先抽样核对。
建议转换后用「字符多重集比对」校验（见主指南 7.11）。
"""
import argparse
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import docx
from docx.oxml.ns import qn
from docx.table import _Cell
from docx.text.paragraph import Paragraph

HEADING_RE = re.compile(r"^Heading\s*(\d)$", re.I)


# ---------- 文本抽取 ----------

def para_text(p):
    """取段落文本，<w:tab>/<w:br> 转空格。"""
    parts = []
    for node in p.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag in (qn("w:tab"), qn("w:br")):
            parts.append(" ")
    return re.sub(r"\s+", " ", "".join(parts)).strip()


def cell_text(tc, table):
    try:
        paras = [para_text(p) for p in _Cell(tc, table).paragraphs]
    except Exception:
        paras = [para_text(p) for p in tc.findall(qn("w:p"))]
    return [t for t in paras if t]


# ---------- 网格还原 ----------

def build_grid(table):
    """返回二维列表，每项为 (list_of_texts, is_vmerge_continue)。"""
    grid = []
    for tr in table._tbl.tr_lst:
        row, col = [], 0
        for tc in tr.tc_lst:
            tcPr = tc.tcPr
            span, vm = 1, None
            if tcPr is not None:
                gs = tcPr.find(qn("w:gridSpan"))
                if gs is not None:
                    try:
                        span = int(gs.get(qn("w:val"), "1"))
                    except ValueError:
                        span = 1
                v = tcPr.find(qn("w:vMerge"))
                if v is not None:
                    vm = "restart" if v.get(qn("w:val")) == "restart" else "continue"

            texts = cell_text(tc, table)
            if vm == "continue":
                for rr in reversed(grid):          # 继承上方同列
                    if col < len(rr):
                        texts = rr[col][0]
                        break
            for _ in range(span):
                row.append((texts, vm == "continue"))
            col += span
        grid.append(row)
    return grid


def collapse_adjacent(cells):
    """折叠行内相邻重复单元格（gridSpan 铺开的值只保留一次）。"""
    out = []
    for c in cells:
        c = c.strip()
        if c and out and out[-1] == c:
            continue
        out.append(c)
    return out


def esc(s):
    return s.replace("|", "\\|")


def render_grid(grid):
    if not grid:
        return []
    ncol = max(len(r) for r in grid)
    if ncol == 0:
        return []

    rows = []
    for r in grid:
        cells = [esc("<br>".join(r[i][0])) if i < len(r) else "" for i in range(ncol)]
        rows.append(collapse_adjacent(cells))

    ncol = max((len(r) for r in rows), default=0)
    if ncol == 0:
        return []
    rows = [r + [""] * (ncol - len(r)) for r in rows]

    # 去尾部全空列
    last = ncol - 1
    while last > 0 and all(not r[last].strip() for r in rows):
        last -= 1
    rows = [r[: last + 1] for r in rows]
    ncol = last + 1

    # 去全空行
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return []

    # 去"空壳行"
    def is_shell(r):
        """判断是否为「空壳行」（Word 里常见的只有首列文字、其余列是序号的行）。

        ★ 必须要求**首列也为空**才算空壳（2026-10-05 修正）：
          原实现只看第 2 列起是否全是数字，于是 `| 采购 | 1 |` 这类
          **首列有内容、其余列是数量**的正常数据行被整行丢弃 —— 真丢字。
          实测一个 2x2 表格（表头「项目/数量」+ 数据行「采购/1」），
          转换后数据行整行消失，`md_lint` 报E-LOSS-01「缺失 3 个字符：1购采」。
        """
        if len(r) < 2:
            return False
        if r[0].strip():          # ★ 首列有内容 ⇒ 是数据行，不是空壳
            return False
        for c in r[1:]:
            c = c.strip()
            if c and not re.fullmatch(r"\d+\s*[、．.,，、]?", c):
                return False
        return True

    rows = [r for r in rows if not is_shell(r)]
    if not rows:
        return []

    # 去相邻完全重复行
    dedup = []
    for r in rows:
        if dedup and dedup[-1] == r:
            continue
        dedup.append(r)

    out = ["| " + " | ".join(dedup[0]) + " |",
           "|" + "|".join([" --- "] * ncol) + "|"]
    out += ["| " + " | ".join(r) + " |" for r in dedup[1:]]
    return out


# ---------- 图片 ----------

def extract_images(document, para, assets_dir, used):
    refs = []
    try:
        blips = para._p.findall(".//" + qn("a:blip"))
    except Exception:
        return refs
    for blip in blips:
        rId = blip.get(qn("r:embed"))
        if not rId:
            continue
        try:
            part = document.part.related_parts[rId]
        except KeyError:
            continue
        ext = os.path.splitext(part.partname)[1] or ".png"
        name = "image_%03d%s" % (len(used) + 1, ext)
        os.makedirs(assets_dir, exist_ok=True)
        with open(os.path.join(assets_dir, name), "wb") as fh:
            fh.write(part.blob)
        used.append(name)
        refs.append("![](%s/%s)" % (os.path.basename(assets_dir), name))
    return refs


# ---------- 主转换 ----------

def convert(docx_path, assets_root=None):
    document = docx.Document(docx_path)
    base = os.path.splitext(os.path.basename(docx_path))[0]
    out_dir = assets_root or os.path.dirname(os.path.abspath(docx_path))
    assets_dir = os.path.join(out_dir, "assets_" + base)
    used = []

    lines = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            p = Paragraph(child, document)
            txt = para_text(child)
            imgs = extract_images(document, p, assets_dir, used)
            if txt:
                try:
                    style = p.style.name
                except Exception:
                    style = None
                m = HEADING_RE.match(style or "")
                if m:
                    lines.append("#" * min(int(m.group(1)), 6) + " " + txt)
                else:
                    lines.append(txt)
            for im in imgs:
                lines += ["", im, ""]
        elif child.tag == qn("w:tbl"):
            table = docx.table.Table(child, document)
            lines.append("")
            lines += render_grid(build_grid(table))
            lines.append("")

    # 首行若为非标题短句（文档标题），提升为 H1
    if lines and lines[0] and not lines[0].startswith(("#", "|")) and len(lines[0]) <= 25:
        lines[0] = "# " + lines[0]

    out, blank = [], 0
    for ln in lines:
        if not ln.strip():
            blank += 1
            if blank > 1:
                continue
        else:
            blank = 0
        out.append(ln)
    return "\n".join(out).strip() + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description="DOCX -> Markdown（修正合并单元格）")
    ap.add_argument("input", help=".docx 文件")
    ap.add_argument("-o", "--output", default=None, help="输出 .md，默认与源文件同名同目录")
    ap.add_argument("--assets-dir", default=None, help="图片输出根目录，默认与 .md 同目录")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.input):
        print("输入不存在: " + args.input)
        return 2

    md = convert(args.input, args.assets_dir)
    dst = args.output or (os.path.splitext(args.input)[0] + ".md")
    d = os.path.dirname(os.path.abspath(dst))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(md)
    print("[完成] %s -> %s （%d 字符）" % (args.input, dst, len(md)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
