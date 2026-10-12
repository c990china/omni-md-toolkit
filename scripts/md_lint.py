# -*- coding: utf-8 -*-
"""MD 自整理体检脚本（只读扫描，绝不修改任何文件）。

按《20260718_全格式转MD并整理指南_MD自整理规范.md》定义的目标态，对转换产出的
Markdown 做体检，输出「规则码 + 行号 + 原文片段」的清单，供人工或 AI 逐项修复。

本脚本只读取文件，不写入、不移动、不删除任何内容。

用法：
  python md_lint.py "a.md"                       # 单文件体检
  python md_lint.py "待整理目录" [-r]             # 目录下所有 .md（-r 递归）
  python md_lint.py "a.md" --source "a.docx"      # 附带无损比对（docx/pdf/txt）
  python md_lint.py "目录" --source-ext docx      # 目录模式下自动按同名源文件比对
  python md_lint.py "a.md" --json                 # 机器可读输出
  python md_lint.py --list-rules                  # 打印规则码表

常用开关：
  --fail-on {error,warning,none}   判定失败的门槛（默认 error）
  --max-per-rule N                 每条规则最多列几条样例（默认 3）
  --ignore CODE,CODE               忽略指定规则码
  --json                           输出 JSON

退出码：
  0 通过      1 存在 E 级问题      2 仅存在 W 级问题（--fail-on warning 时）
  3 运行错误（文件不存在 / 无法解码等）

依赖：仅标准库。--source 为 .docx 时需 python-docx，为 .pdf 时需 PyMuPDF(fitz)；
      缺依赖时该规则自动降级为提示并跳过，不影响其余检查。
"""
import os
import re
import io
import sys
import json
import glob
import shutil
import argparse
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---------------------------------------------------------------- 规则码表

# 码: (级别, 名称)
RULES = {
    # 编码与空白
    "E-ENC-01": ("E", "非 UTF-8 或带 BOM"),
    "W-ENC-02": ("W", "CRLF 与 LF 混用"),
    "W-ENC-03": ("W", "行尾空白（空格/制表符）"),
    "W-ENC-04": ("W", "文件结尾无换行"),
    # 结构与标题
    "E-STR-01": ("E", "H1 缺失或不唯一"),
    "E-STR-02": ("E", "标题层级跳级"),
    "W-STR-03": ("W", "标题以标点结尾"),
    "W-STR-04": ("W", "标题疑似正文粘连"),
    "W-STR-05": ("W", "疑似未标记的编号标题"),
    # 段落
    "W-PAR-01": ("W", "正文超长行且无句末标点"),
    "W-PAR-02": ("W", "孤立残片行"),
    "W-PAR-03": ("W", "连续空行超过 1 行"),
    # 表格
    "E-TBL-01": ("E", "表格块缺分隔行"),
    "E-TBL-02": ("E", "表格列数不一致"),
    "W-TBL-03": ("W", "行内相邻重复单元格"),
    "W-TBL-04": ("W", "相邻完全重复行"),
    "W-TBL-05": ("W", "表格含全空行或尾部空列"),
    "W-TBL-06": ("W", "表格空壳行（仅首列有内容）"),
    "W-TBL-07": ("W", "表格块与上下文之间缺空行"),
    # 图片
    "E-IMG-01": ("E", "图片引用文件不存在"),
    "W-IMG-02": ("W", "图片 alt 为空"),
    "W-IMG-03": ("W", "图片未独立成段"),
    "W-IMG-04": ("W", "图片路径为绝对路径或含反斜杠"),
    "I-IMG-05": ("I", "图片为 base64 内联"),
    # 版面残留
    "W-PG-01": ("W", "独立页码行"),
    "W-PG-02": ("W", "行首页码竖线残留"),
    "W-PG-03": ("W", "标题带页码尾缀"),
    "W-PG-04": ("W", "目录点线残留"),
    "I-PG-05": ("I", "命中广告关键词"),
    # 标点
    "W-PUN-01": ("W", "中文语境半角逗号/句号"),
    "W-PUN-02": ("W", "中文语境半角括号"),
    "I-PUN-03": ("I", "\\uXXXX 转义残留"),
    # 无损
    "E-LOSS-01": ("E", "相对源文件丢字"),
    "I-LOSS-02": ("I", "无损比对跳过（缺依赖或源文件）"),
}

AD_KEYWORDS = [
    "精品课程", "课程售价", "立即购买", "扫码购买", "扫码关注",
    "关注公众号", "二维码", "购买课程", "限时优惠", "加微信",
]

CJK = r"\u4e00-\u9fff"
RE_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*$")
# ★ 图片路径匹配：**括号配平**，而不是「贪婪到最后一个右括号」（2026-10-05 两度修正）。
#   踩坑过程值得记住：
#     ① 原`([^)]+)`：路径含半角括号时（`assets_书名(精校本)/图.png`）在 `精校本)` 处截断，
#        把存在的文件误报成 E-IMG-01（实测 azw3 电子书 3 张图全误报）。
#     ② 改成贪婪 `(.+)`：**一行多图时会把多张图并成一张**——
#        `![a](x.png) ![b](y.png)` 匹配出 src = `x.png) ![b](y.png`，
#        os.path.exists 失败 ⇒报 E 级**阻断**，convert.py 直接拒绝交付（误拦比漏报更糟）。
#     ③ 现用括号配平：内容 = 不含括号的字符，或一段成对的 `(...)`，故
#        既能正确吃掉 `书名(精校本)/图.png`，又能在第一个配平右括号处停下。
#   注：形如`a(b(c)d)e` 的双层嵌套路径不匹配 ⇒ 该图不检查（漏报可接受，误报不可接受）。
RE_LINK_SPAN = re.compile(r"!\[([^\]]*)\]\(((?:[^()]|\([^()]*\))+)\)")
RE_IMG = RE_LINK_SPAN
RE_LIST = re.compile(r"^\s*(?:[-*+]\s+|\d+[.、)）]\s+|[一二三四五六七八九十]+[、.]\s*)")
RE_NUM_HEAD = re.compile(r"^\s*(?:\d+(?:\.\d+)+|[（(][一二三四五六七八九十]+[）)])\s*\S")
RE_SENT_END = "。！？；!?;"

RE_PAGE_ALONE = re.compile(r"^(?:\d{1,4}\s*/\s*\d{1,4}|第\s*\d+\s*页|-\s*\d+\s*-|—\s*\d+\s*—)\s*$")
RE_PAGE_PREFIX = re.compile(r"^\d{1,3}\|")
RE_HEAD_PAGENUM = re.compile(r"^(.+?)\|\s*\d{1,3}\s*[\u2460-\u2469]?\s*$")
RE_TOC_DOTS = re.compile(r"[\.．。｡]{3,}\s*\d+\s*$")
RE_UNICODE_ESC = re.compile(r"\\u[0-9a-fA-F]{4}")

RE_PUN_HALF_COMMA = re.compile(r"[\u4e00-\u9fff]\s*,\s*[\u4e00-\u9fff]")
# 点号后面必须还是中文或已在行尾，才算误用的半角句号；
# 后面接字母/数字的一律放过，否则 xxx.md、toolkit_common.py 会被误判。
RE_PUN_HALF_DOT = re.compile(r"[\u4e00-\u9fff]\.(?=[\u4e00-\u9fff]|\s*$)")
RE_PUN_HALF_PAREN = re.compile(r"[\u4e00-\u9fff]\(|\)[\u4e00-\u9fff]")
RE_BOLD_ONLY = re.compile(r"^\*{1,2}[^*]+\*{1,2}$")
RE_INLINE_CODE = re.compile(r"`[^`]+`")


# ---------------------------------------------------------------- 小工具

def split_cells(line):
    """把 Markdown 表格行拆成单元格列表；非表格行返回 None。"""
    s = line.strip()
    if not s.startswith("|"):
        return None
    s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in re.split(r"(?<!\\)\|", s)]


def is_sep_row(cells):
    return bool(cells) and all(re.fullmatch(r":?-{1,}:?", c) for c in cells if c != "")


def is_blank(s):
    return s.strip() == ""


def looks_like_text(s):
    return bool(re.search(r"[\u4e00-\u9fffA-Za-z0-9]", s))


def clip(s, n=68):
    s = s.replace("\t", " ")
    return s if len(s) <= n else s[: n - 3] + "..."


# ---------------------------------------------------------------- 体检主体

class Report(object):
    def __init__(self, path, source=None):
        self.path = path
        self.source = source
        self.findings = []
        self.stats = {"lines": 0, "tables": 0, "images": 0, "headings": 0}
        self.error = None

    def add(self, code, lineno, snippet, msg=""):
        if code not in RULES:
            return
        level = RULES[code][0]
        self.findings.append({
            "code": code, "level": level, "line": lineno,
            "snippet": clip(snippet), "msg": msg,
        })

    def count(self, level):
        return sum(1 for f in self.findings if f["level"] == level)


def scan_text(text, raw_bytes, rep, opt):
    # ---- 编码与空白（文件级）
    if raw_bytes.startswith(b"\xef\xbb\xbf"):
        rep.add("E-ENC-01", 1, "<文件头>", "存在 UTF-8 BOM，应改为无 BOM")
    crlf = text.count("\r\n")
    lf_only = text.count("\n") - crlf
    if crlf and lf_only:
        rep.add("W-ENC-02", 0, "<全文>", "CRLF %d 行 / LF %d 行混用" % (crlf, lf_only))

    # 行级检查统一按 LF 处理，避免 \r 被误判为行尾空白（行号不变）
    flat = text.replace("\r\n", "\n")
    lines = flat.split("\n")
    rep.stats["lines"] = len(lines)
    if text and not text.endswith("\n"):
        rep.add("W-ENC-04", len(lines), "<文件尾>", "结尾缺少换行")

    def at(i):
        return lines[i - 1] if 1 <= i <= len(lines) else ""

    in_code = False
    code_lines = set()     # 代码围栏内的行号（含围栏行），其中的示例不做检查
    headings = []          # (行号, 级别, 文本)
    table_blocks = []      # (起始行号, [行号...])
    i = 0
    cur_table = None

    while i < len(lines):
        ln = i + 1
        line = lines[i]
        stripped = line.strip()
        # 标点类规则只看正文，不看链接/图片路径（2026-10-05）：
        # 真实文件名允许含半角括号（如 `assets_书名(精校本)/图.png`），
        # 若参与 W-PUN-02 检查会对每张图都误报「中文与半角括号直接相邻」。
        # 用 RE_LINK_SPAN（括号配平）逐个挖掉链接段 —— 不用贪婪 `\]\(.*\)`，
        # 否则一行多图时会把两张图**之间**的正文也一并抹掉，导致标点漏检。
        punct_text = RE_LINK_SPAN.sub("](…)", stripped)

        # 代码块围栏切换
        if stripped.startswith("```") or stripped.startswith("~~~"):
            in_code = not in_code
            code_lines.add(ln)
            if cur_table:
                table_blocks.append(cur_table)
                cur_table = None
            i += 1
            continue
        if in_code:
            code_lines.add(ln)
            i += 1
            continue

        # ---- 连续空行（只在多出来的第一段空行处报一次）
        if is_blank(line) and is_blank(at(ln - 1)):
            # 只在第二段空行处报一次
            if not is_blank(at(ln - 2)):
                rep.add("W-PAR-03", ln, "", "此处起连续空行")

        # ---- 表格块收集
        cells = split_cells(line)
        if cells is not None:
            if cur_table is None:
                cur_table = [ln, []]
            cur_table[1].append(ln)
            i += 1
            continue
        else:
            if cur_table:
                table_blocks.append(cur_table)
                cur_table = None

        # ---- 行尾空白
        if line != line.rstrip():
            rep.add("W-ENC-03", ln, line, "行尾有多余空白")

        if is_blank(line):
            i += 1
            continue

        # ---- 标题
        mh = RE_HEADING.match(line)
        if mh:
            level = len(mh.group(1))
            htext = mh.group(2)
            headings.append((ln, level, htext))
            if htext and htext[-1] in "。，；：、！？,;:":
                rep.add("W-STR-03", ln, line, "标题以标点结尾")
            if len(htext) > 40 and any(c in htext for c in "。；，"):
                rep.add("W-STR-04", ln, line, "标题过长且含句末标点，疑似标题与正文粘连")
            i += 1
            continue

        # ---- 版面残留：页码 / 页眉 / 标题页码 / 目录点线
        if RE_PAGE_ALONE.match(stripped):
            rep.add("W-PG-01", ln, stripped, "独立页码行")
            i += 1
            continue
        if RE_PAGE_PREFIX.match(stripped):
            rep.add("W-PG-02", ln, stripped, "行首形如 `N|` 的页眉页脚残留")
        mhp = RE_HEAD_PAGENUM.match(stripped)
        if mhp and re.search(r"[\u4e00-\u9fff]", mhp.group(1)):
            rep.add("W-PG-03", ln, stripped, "形如 `标题|页码` 的页眉残留")
        if RE_TOC_DOTS.search(stripped):
            rep.add("W-PG-04", ln, stripped, "目录点线残留，应还原为列表或删除")

        # ---- 疑似未标记编号标题
        if (RE_NUM_HEAD.match(stripped) and len(stripped) <= 30
                and not RE_LIST.match(stripped)
                and stripped[-1] not in RE_SENT_END):
            rep.add("W-STR-05", ln, stripped, "形如编号标题但未加 #")

        # ---- 段落类
        if not RE_LIST.match(line) and looks_like_text(stripped):
            if len(stripped) > 200 and not any(c in stripped for c in "。！？；"):
                rep.add("W-PAR-01", ln, stripped, "长度 %d 且无句末标点，疑似一句到底" % len(stripped))
            if (len(stripped) <= 8 and stripped[-1] not in RE_SENT_END + "：:"
                    and not RE_BOLD_ONLY.match(stripped)
                    and is_blank(at(ln - 1)) and is_blank(at(ln + 1))):
                rep.add("W-PAR-02", ln, stripped, "独立成段的短残片，疑似断行残留")

        # ---- 标点（只看正文，不看链接/图片路径，见上方 punct_text）
        if RE_PUN_HALF_COMMA.search(punct_text):
            rep.add("W-PUN-01", ln, stripped, "中文之间出现半角逗号")
        elif RE_PUN_HALF_DOT.search(punct_text):
            rep.add("W-PUN-01", ln, stripped, "中文后出现半角句号")
        if RE_PUN_HALF_PAREN.search(punct_text):
            rep.add("W-PUN-02", ln, stripped, "中文与半角括号直接相邻")
        if RE_UNICODE_ESC.search(stripped):
            rep.add("I-PUN-03", ln, stripped, "存在 \\uXXXX 转义，应还原为原字符")

        # ---- 广告关键词
        for kw in AD_KEYWORDS:
            if kw in stripped:
                rep.add("I-PG-05", ln, stripped, "命中广告关键词「%s」" % kw)
                break

        i += 1

    if cur_table:
        table_blocks.append(cur_table)

    rep.stats["headings"] = len(headings)

    # ---- 标题骨架
    h1 = [h for h in headings if h[1] == 1]
    if len(h1) == 0:
        rep.add("E-STR-01", 0, "<全文>", "缺少一级标题（H1）")
    elif len(h1) > 1:
        rep.add("E-STR-01", h1[1][0], at(h1[1][0]), "H1 不唯一，共 %d 个" % len(h1))
    prev = None
    for ln, level, htext in headings:
        if prev is not None and level > prev + 1:
            rep.add("E-STR-02", ln, at(ln), "H%d 直接跳到 H%d" % (prev, level))
        prev = level

    # ---- 表格体检
    for start, linenos in table_blocks:
        check_table(rep, at, start, linenos)

    # ---- 图片体检
    check_images(rep, flat, lines, os.path.dirname(os.path.abspath(rep.path)), code_lines)

    return rep


def check_table(rep, at, start, linenos):
    rep.stats["tables"] += 1
    rows = [(ln, split_cells(at(ln))) for ln in linenos]
    rows = [(ln, c) for ln, c in rows if c is not None]

    # 上下文空行
    if not is_blank(at(linenos[0] - 1)):
        rep.add("W-TBL-07", linenos[0], at(linenos[0]), "表格首行与上一行之间缺空行")
    if not is_blank(at(linenos[-1] + 1)):
        rep.add("W-TBL-07", linenos[-1], at(linenos[-1]), "表格末行与下一行之间缺空行")

    # 分隔行
    sep_idx = [k for k, (ln, c) in enumerate(rows) if is_sep_row(c)]
    if not sep_idx:
        rep.add("E-TBL-01", start, at(linenos[0]), "表格块（%d 行）缺 `| --- |` 分隔行" % len(rows))
        body = rows
    elif sep_idx[0] != 1:
        rep.add("E-TBL-01", rows[sep_idx[0]][0], at(rows[sep_idx[0]][0]),
                "分隔行不在表格第二行（第 %d 行）" % (sep_idx[0] + 1))
        body = [(ln, c) for k, (ln, c) in enumerate(rows) if k != sep_idx[0]]
    else:
        body = [(ln, c) for k, (ln, c) in enumerate(rows) if k != sep_idx[0]]

    if not body:
        return

    ncols = max((len(c) for ln, c in body), default=0)
    for ln, c in body:
        if len(c) != ncols:
            rep.add("E-TBL-02", ln, at(ln), "列数 %d，与表头 %d 不一致" % (len(c), ncols))

    # 尾部全空列 / 全空行
    if ncols:
        for col in range(ncols - 1, 0, -1):
            if all((col >= len(c)) or (c[col] == "") for ln, c in body):
                rep.add("W-TBL-05", start, at(linenos[0]), "第 %d 列为全空列" % (col + 1))
            else:
                break
    for ln, c in body:
        if not any(x for x in c):
            rep.add("W-TBL-05", ln, at(ln), "全空行")

    # 行内相邻重复单元格（gridSpan 被展开）
    for ln, c in body:
        prev_v = None
        for v in c:
            if v and v == prev_v:
                rep.add("W-TBL-03", ln, at(ln), "相邻重复单元格「%s」，疑似横向合并被展开" % clip(v, 20))
                break
            if v:
                prev_v = v

    # 相邻完全重复行（vMerge 被展开）
    for k in range(1, len(body)):
        if body[k][1] == body[k - 1][1]:
            rep.add("W-TBL-04", body[k][0], at(body[k][0]), "与上一行完全相同，疑似纵向合并被展开")

    # 空壳行：首列也为空、且其余列只剩孤立编号（2026-10-05 与 docx_merge_table 统一口径）
    for ln, c in body:
        if len(c) < 2:
            continue
        # ★ 首列有内容 ⇒ 是正常数据行（如 `| 采购 | 1 |`），不是空壳。
        #   原判定只看「其余列是否全为编号」，会把这类行误报W-TBL-06。
        if c[0].strip():
            continue
        rest = [x for x in c[1:]]
        if rest and all((x == "") or re.fullmatch(r"\d+\s*[、．.,，、]?", x) for x in rest):
            rep.add("W-TBL-06", ln, at(ln), "整行为空壳（仅余孤立编号）")


def in_inline_code(line, col):
    """该列是否落在行内代码反引号内。"""
    for m in RE_INLINE_CODE.finditer(line):
        if m.start() <= col < m.end():
            return True
    return False


def check_images(rep, text, lines, base_dir, skip_lines=None):
    """skip_lines：代码围栏内的行号集合，其中的图片示例不算真实引用。

    另外两类也不是真实引用，一并跳过：
      - 表格行（正误对照表里必然出现反例写法）
      - 行内代码反引号内的示例
    """
    skip_lines = skip_lines or set()
    for m in RE_IMG.finditer(text):
        lineno = text[: m.start()].count("\n") + 1
        if lineno in skip_lines:
            continue
        line = lines[lineno - 1] if lineno - 1 < len(lines) else ""
        if line.lstrip().startswith("|"):
            continue
        line_start = text.rfind("\n", 0, m.start()) + 1
        if in_inline_code(line, m.start() - line_start):
            continue
        alt, src = m.group(1), m.group(2).strip()
        rep.stats["images"] += 1

        if src.startswith("data:"):
            rep.add("I-IMG-05", lineno, clip(m.group(0)), "base64 内联图片，应落盘为文件")
            continue
        if src.startswith("http://") or src.startswith("https://"):
            pass
        else:
            if os.path.isabs(src) or "\\" in src:
                rep.add("W-IMG-04", lineno, m.group(0), "应使用相对路径且以 / 分隔")
                probe = src.replace("\\", "/")
            else:
                probe = src
            full = os.path.join(base_dir, probe)
            if not os.path.exists(full):
                rep.add("E-IMG-01", lineno, m.group(0), "引用文件不存在：%s" % probe)

        if not alt.strip():
            rep.add("W-IMG-02", lineno, m.group(0), "alt 为空，应写图意 + 图号")
        if line.strip() != m.group(0).strip():
            rep.add("W-IMG-03", lineno, line, "图片未独立成段（与文字同行）")


# ---------------------------------------------------------------- 无损比对

def norm_for_loss(t):
    t = re.sub(r"<br\s*/?>", "", t)
    t = t.replace("|", "").replace("#", "").replace("\\", "")
    t = re.sub(r"!\[.*?\]\(.*?\)", "", t)
    t = re.sub(r"<!--.*?-->", "", t)
    # ★ 行首列表编号的标点属于**版式规范化**，不是内容：
    #   自整理会把 `1、甲` 规范成 `1.甲`、把 `1）甲` 规范成 `1.甲`。若不在这里抹平，
    #   无损比对会误报「相对源文件缺失 1 个字符」（E-LOSS-01），把整批任务判成失败。
    #   做法：把行首「数字/中文数字 + 、.．。）)」与「（数字）」统一成 "."，两侧一致后再比对。
    t = re.sub(r"(?m)^[ \t]*(?:\d{1,4}|[一二三四五六七八九十]{1,3})[.、．。)）\]](?=\s*\S)", ".", t)
    t = re.sub(r"(?m)^[ \t]*[（(](?:\d{1,4}|[一二三四五六七八九十]{1,3})[）)](?=\s*\S)", ".", t)
    #★ 省略号写法差异不算丢字（2026-10-05）：原文排版成半角多点（`...`、个别书两点的 `..`），
    #   自整理按规范统一成标准省略号 `……`。字符集不同（`...` 是 3 个 `.`，`……` 是 2 个 `…`），
    #   不抹平就会误报 E-LOSS-01「缺失 3 个字符」（实测某书原文 3 处 `...`）。
    #   ★ 只把「半角点串」折成 1 个 `…`，**不能**把 `……` 也折叠
    #     （`……`→`…` 会凭空少 1 个字符，制造新的假阳性）。
    #   ★ 两点 `..` 也折：`..`→`…` 与 `..`→`……` 字符数不同，会误报 1 个字符。
    t = re.sub(r"\.{2,}", "…", t)
    # ★ 全角/半角标点等价（2026-10-05）：自整理会把中文语境的 `(),` 规范成全角 `（），`，
    #   字符集不同（各 1 个字符但形态不同）⇒ 不抹平会误报 E-LOSS-01「缺失 2 个字符」。
    #   实测一个含 `(半角)` 的 txt 在 HEAD 版报缺失 7、新版报缺失 4，其中 2 个正是括号。
    #   这是**版式规范化**而非丢字，与列表编号归一同理。
    t = t.translate(str.maketrans("（）", "()"))
    return re.sub(r"\s+", "", t)


def source_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md", ".text"):
        with open(path, encoding="utf-8", errors="ignore") as f:
            return f.read(), None
    if ext in (".azw3", ".mobi", ".kf8", ".epub"):
        # ★ 电子书无损比对（2026-10-05 新增）：解包后取**纯正文文本**作为比对基准。
        #   此前 `.azw3/.epub` 一律走末尾的「不支持」分支 ⇒ convert.py 传 source=None
        #   ⇒ **电子书只过格式门禁、丢字完全不报警**（与 FAQ 10/11 同类的「静默不检查」）。
        #   注意：基准是**未自整理的抽取原文**，所以成品里的版式规范化
        #   （编号 `1、`→`1.`、标点全角化等）由 norm_for_loss 统一归一，两侧对称。
        try:
            import importlib.util
            here = os.path.dirname(os.path.abspath(__file__))
            spec = importlib.util.spec_from_file_location(
                "kindle_to_md_lossless",
                os.path.join(here, "kindle_to_md.py"))
            k = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(k)
            import tempfile
            tmp = tempfile.mkdtemp(prefix="lossless_")
            try:
                md = k.extract(path, os.path.join(tmp, "assets"), "src")
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            return md, None
        except Exception as e:
            return None, "电子书无损比对失败（缺 mobi/bs4 或解包错误）：%s" % e
    if ext == ".docx":
        try:
            import docx
            from docx.oxml.ns import qn
        except Exception:
            return None, "缺 python-docx，无法读取 docx 源文件"
        d = docx.Document(path)
        parts = []
        for p in d.element.body.iter(qn("w:p")):
            seg = []
            for node in p.iter():
                if node.tag == qn("w:t"):
                    seg.append(node.text or "")
                elif node.tag in (qn("w:tab"), qn("w:br")):
                    seg.append(" ")
            parts.append("".join(seg))
        # ★ 段落之间必须保留换行：`norm_for_loss()` 的行首编号归一化靠 `(?m)^` 锚点，
        #   若把段落拼成一行，原文里的 `1、` 就永远匹配不到，会造成 E-LOSS-01 误报
        #   （成品 md 有换行、会被抹平，两侧不对称）。空白最终会被 norm 掉，不影响字符集比对。
        return "\n".join(parts), None
    if ext == ".pdf":
        try:
            import fitz
        except Exception:
            return None, "缺 PyMuPDF(fitz)，无法读取 pdf 源文件"
        doc = fitz.open(path)
        try:
            return "".join(doc.load_page(i).get_text() for i in range(doc.page_count)), None
        finally:
            doc.close()
    return None, "不支持的源文件类型：%s（支持 azw3/mobi/kf8/epub/docx/pdf/txt/md）" % ext


def check_lossless(rep, opt):
    src = rep.source
    if not src:
        return
    if not os.path.exists(src):
        rep.add("I-LOSS-02", 0, src, "源文件不存在，跳过无损比对")
        return
    stext, err = source_text(src)
    if err:
        rep.add("I-LOSS-02", 0, src, err)
        return
    try:
        with open(rep.path, encoding="utf-8", errors="ignore") as f:
            mtext = f.read()
    except Exception as e:
        rep.add("I-LOSS-02", 0, rep.path, "MD 读取失败：%s" % e)
        return
    cs = Counter(norm_for_loss(stext))
    cd = Counter(norm_for_loss(mtext))
    missing = cs - cd
    miss = sum(missing.values())
    if miss > 0:
        sample = "".join(sorted(missing.elements()))[:60]
        rep.add("E-LOSS-01", 0, "<全文>",
                "相对源文件缺失 %d 个字符；样例：%s" % (miss, sample))


# ---------------------------------------------------------------- 驱动

def lint_file(path, opt):
    rep = Report(path, None)
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except Exception as e:
        rep.error = "读取失败：%s" % e
        return rep
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError as e:
        rep.add("E-ENC-01", 0, "<全文>", "非 UTF-8 编码：%s" % e)
        try:
            text = raw.decode("utf-8", errors="replace")
        except Exception:
            rep.error = "无法解码"
            return rep
    else:
        text = raw.decode("utf-8")

    # 无损比对：单文件用 --source；目录模式用 --source-ext 按同名配对
    src = None
    if opt.source_ext:
        stem = os.path.splitext(os.path.basename(path))[0]
        base = opt.source_dir or os.path.dirname(os.path.abspath(path))
        cand = os.path.join(base, stem + "." + opt.source_ext.lstrip("."))
        src = cand if os.path.exists(cand) else None
    elif opt.source:
        src = opt.source
    if src:
        rep.source = src
        check_lossless(rep, opt)

    try:
        scan_text(text, raw, rep, opt)
    except Exception as e:
        rep.error = "扫描异常：%s" % e
    return rep


def collect_targets(target, opt):
    if os.path.isfile(target):
        return [target]
    pat = os.path.join(target, "**", "*.md") if opt.recursive else os.path.join(target, "*.md")
    return sorted(glob.glob(pat, recursive=opt.recursive))


def print_report(rep, opt):
    print("=" * 66)
    print("文件: %s" % rep.path)
    if rep.source and os.path.exists(rep.source):
        print("源文件: %s" % rep.source)
    if rep.error:
        print("!! %s" % rep.error)
        return
    st = rep.stats
    print("行数 %d / 标题 %d / 表格块 %d / 图片引用 %d"
          % (st["lines"], st["headings"], st["tables"], st["images"]))

    by_code = {}
    for f in rep.findings:
        by_code.setdefault(f["code"], []).append(f)

    for code in sorted(by_code):
        fs = by_code[code]
        level, name = RULES[code]
        print("\n[%s] %s  %s  x%d" % (level, code, name, len(fs)))
        for f in fs[: opt.max_per_rule]:
            loc = "L%d" % f["line"] if f["line"] else "--"
            print("    %-8s %s" % (loc, f["snippet"]))
            if f["msg"]:
                print("             -> %s" % f["msg"])
        if len(fs) > opt.max_per_rule:
            print("    ... 另有 %d 条同类" % (len(fs) - opt.max_per_rule))

    e, w, i = rep.count("E"), rep.count("W"), rep.count("I")
    print("\n-- 汇总 --  E %d   W %d   I %d" % (e, w, i))
    if e:
        print("结论: 未通过（存在 E 级问题，阻断交付）")
    elif w:
        print("结论: 有条件通过（W 级需人工确认）")
    else:
        print("结论: 通过")


def print_rules():
    print("%-12s %-4s %s" % ("规则码", "级别", "含义"))
    print("-" * 66)
    for code in sorted(RULES):
        level, name = RULES[code]
        print("%-12s %-4s %s" % (code, level, name))


def main():
    ap = argparse.ArgumentParser(
        description="MD 自整理体检（只读扫描，绝不修改文件）")
    ap.add_argument("target", nargs="?", help="待体检的 .md 文件或目录")
    ap.add_argument("-r", "--recursive", action="store_true", help="目录模式下递归子目录")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--source", default=None,
                    help="源文件路径（单文件模式做无损比对）；目录模式配合 --source-ext 使用")
    ap.add_argument("--source-ext", default=None,
                    help="目录模式：按同名同目录的该扩展名源文件做无损比对，如 docx")
    ap.add_argument("--source-dir", default=None,
                    help="目录模式：源文件所在目录（默认与 MD 同目录）")
    ap.add_argument("--fail-on", choices=["error", "warning", "none"], default="error",
                    help="判定失败的门槛（默认 error）")
    ap.add_argument("--max-per-rule", type=int, default=3, help="每条规则最多列几条样例")
    ap.add_argument("--ignore", default="", help="忽略的规则码，逗号分隔")
    ap.add_argument("--list-rules", action="store_true", help="打印规则码表后退出")
    a = ap.parse_args()

    if a.list_rules:
        print_rules()
        return 0
    if not a.target:
        ap.error("缺少待体检目标；用 --list-rules 查看规则码表")

    opt = a

    files = collect_targets(a.target, opt)
    if not files:
        sys.stderr.write("[错误] 未找到任何 .md：%s\n" % a.target)
        return 3

    ignored = set(x.strip().upper() for x in a.ignore.split(",") if x.strip())

    # --source 只在单文件模式直接生效；多文件模式请用 --source-ext 逐个配对
    if len(files) > 1 and a.source and not a.source_ext:
        sys.stderr.write("[提示] 多文件模式下 --source 已忽略，改用 --source-ext 按同名配对。\n")
        opt.source = None

    reports = []
    for f in files:
        rep = lint_file(f, opt)
        if ignored:
            rep.findings = [x for x in rep.findings if x["code"] not in ignored]
        reports.append(rep)

    if a.json:
        out = []
        for rep in reports:
            out.append({
                "file": rep.path,
                "source": rep.source if rep.source and os.path.exists(rep.source) else None,
                "error": rep.error,
                "stats": rep.stats,
                "findings": rep.findings,
                "summary": {"E": rep.count("E"), "W": rep.count("W"), "I": rep.count("I")},
            })
        sys.stdout.write(json.dumps(out, ensure_ascii=False, indent=2) + "\n")
    else:
        for rep in reports:
            print_report(rep, opt)
        if len(reports) > 1:
            te = sum(r.count("E") for r in reports)
            tw = sum(r.count("W") for r in reports)
            print("\n" + "=" * 66)
            print("总计 %d 个文件：E %d   W %d" % (len(reports), te, tw))

    tot_e = sum(r.count("E") for r in reports)
    tot_w = sum(r.count("W") for r in reports)
    if any(r.error for r in reports):
        return 3
    if tot_e:
        return 1
    if tot_w and a.fail_on == "warning":
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
