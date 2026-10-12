# -*- coding: utf-8 -*-
"""Markdown 表格事后修复（把「已转好但坏掉」的表格修回合法 GFM 表格）。

适用场景：源文档里 Word 的合并单元格被展开成重复值，转出的 MD 表格因此坏掉：
  - `w:gridSpan`（横向合并 N 列）→ 同一个值在相邻列重复 N 次（`姓名 | 代立信 | 代立信`）
  - `w:vMerge`（纵向合并）→ 同一行在相邻行重复出现多次
  - 且缺少 `| --- |` 表头分隔行 → 根本不是合法 Markdown 表格

修复动作（按顺序）：
  1. 折叠**行内相邻重复单元格**（空单元格不折叠）
  2. 去掉行首/行尾空单元格
  3. 删除全空行
  4. 删除**连续重复的整行**
  5. 列数对齐（不足补空）
  6. 补齐表头分隔行，使其成为合法 GFM 表格

⚠️ 注意：本工具是**事后修复器**，默认**只预览不写盘**（需显式加 --apply）。
   它假定输入表格是「被合并单元格污染」的坏表；若你的表格本身正常、
   只是相邻两列**恰好**是同一个值，折叠会误伤 —— 此时加 --no-collapse 跳过第 1 步。

用法：
  python scripts/fix_tables.py "x.md"            # 预览
  python scripts/fix_tables.py "x.md" --apply    # 写入
  python scripts/fix_tables.py "某个目录" -r --apply
"""
import os
import sys
import glob
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# ---------------------------------------------------------------- 单行 / 单元格

def split_cells(line):
    """`| a | b |` -> ['a','b']；非表格行返回 None。"""
    s = line.strip()
    if not s.startswith("|"):
        return None
    s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def collapse_row(cells):
    """折叠相邻重复单元格（空单元格不参与折叠）。"""
    out = []
    for c in cells:
        if c and out and out[-1] == c:
            continue
        out.append(c)
    return out


def strip_empty_edges(cells):
    while cells and not cells[0]:
        cells.pop(0)
    while cells and not cells[-1]:
        cells.pop()
    return cells


def is_sep(cells):
    """是否为 `| --- |` 之类的分隔行。"""
    return bool(cells) and all(set(c) <= set("-: ") and c for c in cells)


# ---------------------------------------------------------------- 单表 / 全文

def process_table(rows, log, collapse=True):
    """rows: list[list[str]]。返回整理后的 rows（含分隔行）；空则返回 []。"""
    rows = [r for r in rows if not is_sep(r)]
    if not rows:
        return []

    if collapse:
        collapsed = []
        for r in rows:
            c = strip_empty_edges(collapse_row(r))
            collapsed.append(c)
        rows = collapsed
    else:
        rows = [strip_empty_edges(list(r)) for r in rows]

    rows = [r for r in rows if any(x for x in r)]
    if not rows:
        return []

    dedup = []
    for r in rows:
        if dedup and dedup[-1] == r:
            log.append("删重复行: " + " | ".join(r)[:60])
            continue
        dedup.append(r)
    if not dedup:
        return []

    ncol = max(len(r) for r in dedup)
    for r in dedup:
        while len(r) < ncol:
            r.append("")

    return [dedup[0], ["---"] * ncol] + dedup[1:]


def process_text(text, log, collapse=True):
    """全文处理：逐行扫描，遇到连续 `|` 行即视为一张表；代码围栏内不动。"""
    lines = text.split("\n")
    out = []
    buf = []
    in_code = False

    def flush():
        nonlocal buf
        if buf:
            tbl = process_table(buf, log, collapse)
            for r in tbl:
                out.append("| " + " | ".join(r) + " |")
            if not tbl:
                log.append("整表被清空（原 %d 行）" % len(buf))
            buf = []

    for ln in lines:
        if ln.strip().startswith("```"):
            in_code = not in_code
            flush()
            out.append(ln)
            continue
        if in_code:
            out.append(ln)
            continue
        cells = split_cells(ln)
        if cells is not None:
            buf.append(cells)
        else:
            flush()
            out.append(ln)
    flush()
    return "\n".join(out)


# ---------------------------------------------------------------- CLI

def _iter_targets(path, recursive):
    if os.path.isfile(path):
        return [path]
    if os.path.isdir(path):
        pat = "**/*.md" if recursive else "*.md"
        return sorted(glob.glob(os.path.join(path, pat), recursive=recursive))
    return []


def fix_file(path, apply=False, collapse=True):
    """返回 (changed: bool, log: list[str])。"""
    raw = open(path, encoding="utf-8-sig").read()
    log = []
    new = process_text(raw, log, collapse)
    changed = new != raw
    if changed and apply:
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(new)
    return changed, log


def main():
    ap = argparse.ArgumentParser(description="Markdown 表格事后修复（把坏表修回合法 GFM 表格）")
    ap.add_argument("path", help="单个 .md 文件，或包含 .md 的目录")
    ap.add_argument("-r", "--recursive", action="store_true", help="目录模式递归子目录")
    ap.add_argument("--apply", action="store_true", help="真正写入（默认只预览）")
    ap.add_argument("--no-collapse", action="store_true",
                    help="跳过「折叠相邻重复单元格」这一步（表格本身正常、只是相邻列同值时用）")
    a = ap.parse_args()

    targets = _iter_targets(a.path, a.recursive)
    if not targets:
        sys.stderr.write("[错误] 没找到 .md 文件：%s\n" % a.path)
        return 1

    changed_n = 0
    for f in targets:
        try:
            changed, log = fix_file(f, a.apply, not a.no_collapse)
        except Exception as e:  # noqa: BLE001
            sys.stderr.write("[错误] %s -> %s: %s\n" % (f, type(e).__name__, e))
            continue
        tag = ("已修复" + ("（已写入）" if a.apply else "（预览，未写入）")) if changed else "无变化"
        if changed:
            changed_n += 1
        sys.stderr.write("[%s] %s\n" % (tag, os.path.basename(f)))
        for l in log[:8]:
            sys.stderr.write("   " + l + "\n")
        if len(log) > 8:
            sys.stderr.write("   ...（共 %d 处改动）\n" % len(log))

    sys.stderr.write("\n[汇总] 扫描 %d 个文件，%d 个需要修复。%s\n"
                     % (len(targets), changed_n,
                        "" if a.apply else "当前为预览模式，加 --apply 才会写入。"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
