# -*- coding: utf-8 -*-
"""Step 2: Markdown 整理脚本（OCR 后处理）
功能：标点纠错 | 表格还原 | 章节归并 | 页码清理 | 内容去重 | 附件提取 | 编号标题归一
用法：
  python md_reorganizer.py "xxx.md"                  # 默认**原地更新**该文件
  python md_reorganizer.py "xxx.md" --output "out.md"
  python md_reorganizer.py "xxx.md" --no-dedup
  python md_reorganizer.py "xxx.md" --proj-md        # 套用项目规范：去行首 > 、'- '列表转 1）2）3）4）全角编号
"""
import sys, io, os, argparse, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ========== 标点纠错映射 ==========
# 修复：原版 (r'\.', '。') 和 (r',', '，') 对全文无差别替换，
# 导致编号 3.4.1→3。4。1、URL、文件扩展名全损坏。
# 改为上下文限定：仅在中文语境中替换。
PUNCT_REPLACEMENTS = [
    # ★ 先处理省略号，再处理句号（2026-10-05 修正顺序）：
    #   规则 `([一-鿿])\.+` 会把「中文 + 三个半角点」整个吃成 `中文。`，
    #   而 `...` 在许多书的排版里就是**合法省略号** ⇒ 静默丢字
    #   （实测某书原文 3 处 `...`，成品变成 `。`，字符级比对报 E-LOSS-01 缺失 3 字符）。
    #   ★ 判定只要求**左侧紧邻文字**——省略号常出现在**段末**（右侧是换行），
    #     若要求两侧都有文字会漏判（实测 2 处段末省略号被吞成句号）。
    (r'([一-鿿\w])\.{2,}', r'\1……'),                    # 文字后的多点 → 标准省略号（段末亦可）
    (r'([一-鿿])\.(?![.\d])', r'\1。'),                    # 中文后的孤立单点 → 句号
    # 中文逗号：中文字 + 半角逗号 + 中文字 → 全角逗号
    (r'([一-鿿]),([一-鿿])', r'\1，\2'),
    # 中文括号：中文字 + 半角括号 → 全角括号（仅在中文字紧邻时）
    (r'([一-鿿])\(', r'\1（'),
    (r'\)([一-鿿])', r'）\1'),
    # 日语辅助符号清除
    (r'[々〇〻]', ''),
    # OCR 常见误识别
    # ★ 多点 → 句号的规则已上移到 PUNCT_REPLACEMENTS 开头（须先于单点句号规则，
    #   否则 `([一-鿿])\.+` 会把合法省略号 `...` 吞掉 —— 静默丢字，2026-10-05 修正）。
    (r'(?<![一-鿿\w])\.{2,}', '。'),  # 行首/行尾的孤立多点 → 句号（OCR 噪声）
    (r'[,，。]+([）)）])', r'\1'),  # 标点后多余逗号
    # 编号格式纠错
    (r'^(\d+)[、，](?![\d])', r'\1. '),   # `1、` → `1. `
    (r'^(\d+)[．\.。](?![A-Z\d])', r'\1. '),  # `1．` → `1. `
    # 分项编号（续）
    (r'^(\d+)\s*[．、](?=\S)', r'\1. '),  # `1、xxx` → `1. xxx`
    # 冒号对齐项
    (r'^\s*(\S[^\n：]*)\s+:\s*', r'\1：'),   # 末尾空格+冒号 → 中文冒号
    (r'^\s*(\S[^\n：]*[^\s　:：])\s{2,}([^\n]+)$', r'\1\2'),  # 多空格压缩
]

# ========== 正则 ==========
RE_PAGE_NUM    = re.compile(r'^(\d+)/(\d+)$')              # 3/15
RE_ATTACH      = re.compile(r'^附件\s*[1１2２两][：:．.]?\s*')
RE_DUP_HEAD    = re.compile(r'^## 第\s*\d+\s*页\s*$')

# Markdown 行内链接 / 图片：路径用**括号配平**匹配（2026-10-05）。
# 内容 = 不含括号的字符，或一段成对的 `(...)`。这样既能正确吃掉
# `assets_书名(精校本)/图.png` 这类含半角括号的真实路径，又能在第一个配平
# 右括号处停下 —— 一行多图（`![a](x.png) ![b](y.png)`）不会被并成一张。
RE_LINK_SPAN   = re.compile(r'!\[([^\]]*)\]\(((?:[^()]|\([^()]*\))+)\)')

# ========== 表格检测（基于列对齐特征）============

def detect_table(lines):
    """
    识别疑似表格行：包含多列对齐的文字（以空白或分隔符分隔）
    返回：(table_rows, non_table_lines)
    """
    table_rows = []
    non_table  = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        parts = re.split(r'\s{2,}', stripped)
        if len(parts) >= 3 and all(len(p) < 25 for p in parts):
            table_rows.append(parts)
        elif len(parts) == 2:
            if len(parts[0]) < 12 and len(parts[1]) > 10:
                table_rows.append(parts)
            else:
                non_table.append(line)
        else:
            non_table.append(line)

    return table_rows, non_table

def try_build_md_table(rows):
    """
    将检测到的行还原为 Markdown 表格
    rows: list of [col1, col2, ...]
    """
    if not rows:
        return []
    col_counts = [(i, len(r)) for i, r in enumerate(rows)]
    header_idx = max(col_counts, key=lambda x: x[1])[0]
    header = rows[header_idx]
    n_cols = len(header)

    valid = [r for r in rows if len(r) >= n_cols - 1 and len(r) <= n_cols + 1]
    if len(valid) < 2:
        return []

    def pad(row):
        while len(row) < n_cols:
            row.append('')
        return row[:n_cols]

    md_lines = []
    md_lines.append('| ' + ' | '.join(pad(list(header))) + ' |')
    md_lines.append('|' + '|'.join(['---'] * n_cols) + '|')
    for r in rows[header_idx+1:]:
        md_lines.append('| ' + ' | '.join(pad(list(r))) + ' |')

    return md_lines

# ========== 核心整理逻辑 ==========

def clean_line(line):
    """单行纠错"""
    if not line.strip():
        return None

    # [Bug①修复 20260906] 表格行保护：| 开头行直接返回，不跑标点压缩规则
    # 原因：多空格压缩正则会把 `| X |  |` 压成 `| X ||`，损坏 GFM 表格
    if line.strip().startswith('|'):
        return line

    # [Bug②修复 20261005] 行内链接 / 图片目标保护：`](...)` 内的路径不跑标点规则
    # 原因：中文括号全角化规则 `([一-鿿])\(` → `\1（` 会把
    #   `![图](assets_目的与利润(精校本)/x.png)` 改成 `（精校本)`，
    # 路径失配 ⇒ md_lint 报E-IMG-01「图片引用文件不存在」，图片全点不开。
    #   凡路径里带中文括号的 assets 目录都会中招（实测 azw3 电子书）。
    # 处理：把链接段替换为占位符 → 跑规则 → 原样还原（幂等）。
    # ★ 用**括号配平**匹配 `\]\(((?:[^()]|\([^()]*\))+)\)`：
    #   不用 `[^)]*`——会被路径自身的 `(` 提前截断（实测踩过）；
    #   也不用贪婪 `.*`——一行多图时会吞掉两张图之间的正文标点，导致标点漏检。
    spans = []

    def _stash(m):
        spans.append(m.group(0))
        return "\x00LNK%d\x00" % (len(spans) - 1)

    stashed = RE_LINK_SPAN.sub(_stash, line)
    if stashed != line:
        line = stashed

    m = RE_PAGE_NUM.match(line.strip())
    if m and 1 <= int(m.group(1)) <= 1000:
        return None

    if re.match(r'^第\s*\d+\s*页\s*$', line.strip()):
        return None

    result = line
    for pattern, replacement in PUNCT_REPLACEMENTS:
        result = re.sub(pattern, replacement, result)
    for i, seg in enumerate(spans):
        result = result.replace("\x00LNK%d\x00" % i, seg)
    return result

# 章节标题关键词黑名单
_TITLE_BLACKLIST = {
    '现场负责人', '熟练工', '初级工', '人天', '备注', '序号',
    '数量', '单位', '采购分项', '人员类型', '项目类别',
    '姓名', '职称', '岗位', '职责', '合计',
}


def _load_extra_title_blacklist():
    """读取**用户可维护**的标题黑名单文件（每行一条，# 开头为注释）。

    用途：`classify_section` 靠「行长 + 标点 + 关键词」猜标题，会把某些**正文短句**
    误升为标题。实测某书正文「现在是时候把大象提出来了」（原文是段落）被提升为 `##`，
    而电子书抽取端已正确输出了全部真标题 —— 这类误判只能靠黑名单压制。
    若不建本文件，用户每次重跑都要手工改一遍，改动不可持续。

    文件位置（按优先级取第一个存在的）：
      1. 环境变量 `TITLE_BLACKLIST` 指定的路径；
      2. 本脚本同目录的 `标题黑名单.txt`；
      3. 仓库根目录（本脚本上一级）的 `标题黑名单.txt`。
    都不存在时静默跳过。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.environ.get('TITLE_BLACKLIST', ''),
        os.path.join(here, '标题黑名单.txt'),
        os.path.join(os.path.dirname(here), '标题黑名单.txt'),
    ]
    out = set()
    for path in candidates:
        if not path or not os.path.isfile(path):
            continue
        try:
            with open(path, encoding='utf-8') as f:
                for line in f:
                    s = line.strip()
                    if s and not s.startswith('#'):
                        out.add(s)
        except Exception:
            pass
        break
    return out


_EXTRA_TITLE_BLACKLIST = _load_extra_title_blacklist()

def classify_section(text):
    """
    判断文本是否属于章节标题
    返回：(level, clean_text) 或 (0, text)
    level: 1=一级 ##, 2=二级 ###, 3=三级 ####
    """
    t = text.strip()
    if not t or len(t) > 100:
        return 0, text

    # ★ 用户可维护黑名单（标题黑名单.txt）优先于一切启发式判定。
    #   场景：抽取端已正确给出全部真标题，但 `classify_section` 仍会把某些**正文短句**
    #   误升为标题（实测「现在是时候把大象提出来了」——原文是段落）。
    #   这类误判只能靠人工登记压制，否则每次重跑都要手工改一遍。
    if t in _EXTRA_TITLE_BLACKLIST:
        return 0, text

    if re.match(r'^\d+$', t):
        return 0, t
    if t in _TITLE_BLACKLIST:
        return 0, t

    # 模式1: X.X.X 结构（1.1, 1.2.3 等）
    m = re.match(r'^(\d+(?:\.\d+)+)[．.、\s]+(.+)$', t)
    if m:
        num     = m.group(1)
        content = m.group(2).strip()
        if content in _TITLE_BLACKLIST or len(content) < 2:
            return 0, t
        dot_count = num.count('.')
        level = min(dot_count + 1, 3)
        return level, t

    # 模式2: 第X章 / 第X节
    m2 = re.match(r'^第[一二三四五六七八九十百千万\d]+[章节篇部题项][一-鿿]*$', t)
    if m2:
        return 1, t

    # 模式3: 短行无标点
    if 4 <= len(t) < 30 and not re.search(r'[。！？；]$', t) \
       and not re.search(r'[，、：]', t) \
       and not re.match(r'^(图\d|目前|近年来|当前|截至|随着|此外|同时|一是|二是|三是)', t) \
       and re.match(r'^[一-鿿]', t):
        return 1, t

    return 0, t

def deduplicate_headings(headings_seen, level, text):
    """章节去重"""
    key = (level, text)
    if key in headings_seen:
        return False, headings_seen
    headings_seen.add(key)
    return True, headings_seen

# ========== 编号条款标题归一 ==========

RE_HEAD_DETECT = re.compile(r'^(\d+(?:\.\d+)*)\.?\s')
RE_TOC_FILLER = re.compile(r'[\.．。｡]{3,}\s*\d+\s*$')
RE_UNIT = re.compile(r'^\d+(?:\.\d+)*\s*(?:[米厘米毫米千米公里kcmMm²m³吨千克公斤克kgtg台个件项张页套倍层条人次元年月日%％°]+)$')

def normalize_numbered_headings(md_text):
    """按「编号段数」统一标题级别（确定性后处理，覆盖扫描/文本两条路线）。

    级别 = 编号段数 + 1（段数 = 用 '.' 分隔的段数）：
      - `1` / `1.` / `5.`       -> 1 段 -> `##`（h2，章/节）
      - `1.1` / `5.1` / `3.1`   -> 2 段 -> `###`（h3，条）
      - `1.1.1` / `3.4.1`       -> 3 段 -> `####`（h4，款）
    """
    out = []
    for line in md_text.split('\n'):
        s = line.strip()
        if not s:
            out.append(line); continue
        if s.startswith('-') or s.startswith('*'):
            out.append(line); continue
        content = s.lstrip('#').strip() if s.startswith('#') else s
        m = RE_HEAD_DETECT.match(content)
        if not m:
            out.append(line); continue
        if RE_TOC_FILLER.search(s):
            out.append(line); continue
        if RE_UNIT.match(content):
            out.append(line); continue
        # ★ 2026-10-05：此处必须查标题黑名单。
        #   本函数用 `RE_HEAD_DETECT` 正则独立判定编号标题，**不经过 `classify_section`**，
        #   导致用户登记在 `标题黑名单.txt` 的条目对它无效。
        #   实测：某书正文里的编号列表项 `1.谁能帮助你最多？…` 被 `split_body_sentences`
        #   拆成独立短行后，由本函数升成 `## 1. 谁能帮助你最多？`（多出 5 个假章节）。
        #   黑名单对「reorganize 路径」有效、对本函数无效 ⇒ 用户会以为登记没生效。
        #   故在此补查，保持「黑名单优先于一切标题判定」这一统一语义。
        if _EXTRA_TITLE_BLACKLIST and content in _EXTRA_TITLE_BLACKLIST:
            out.append(line); continue
        num = m.group(1)
        depth = len(num.split('.'))
        level = min(depth + 1, 6)
        if not s.startswith('#') and len(content) > 40:
            out.append(line); continue
        out.append('#' * level + ' ' + content)
    return '\n'.join(out)


# ========== 主整理函数 ==========

def full_reorganize(text, dedup=True, verbose=False):
    """**完整**自整理流水线（CLI 步骤 0/7 ~ 7/7 的库函数版）。

    ★ 2026-10-05 新增：此前这20 余步只写在 CLI `main()` 里，而 `convert.py::postprocess()`
      只调了其中的 `reorganize()` + `normalize_blank_lines()` ⇒
      **同一条链路因入口不同而产出不同**：实测同一 azw3，
      单文件入口得719 行 / 20 标题 / W0，批量流水线得 2059 行 / 25 标题 / W5
      （目录项被误升为 5 个 `##`）。这正是本文件 SKILL 踩坑 #14 警告的
      「同一份自整理逻辑不要在 CLI 与库两条路径各写一遍，否则必然漂移」。
      ⇒ 现在两条路径统一调用本函数，**行为一致**。
    """
    def p(msg):
        if verbose:
            print("  " + msg, flush=True)

    # Step 0
    p("步骤 0/7: 清理页眉页脚残留 + 封面残片...")
    text = strip_pdf_headers_footers(text)
    text = _fix_cover_fragments(text)
    # Step 1
    p("步骤 1/7: 合并断裂标题...")
    text = merge_broken_headings(text)
    p("步骤 1b/7: 正文长行内标题粘连拆分...")
    text = split_runon_body(text)
    # Step 2
    p("步骤 2/7: 拆分标题正文粘连...")
    text = split_heading_body(text)
    # Step 3
    p("步骤 3/7: 加粗子标题转标题...")
    text = convert_bold_subheadings(text)
    # Step 4
    p("步骤 4/7: 标点纠错 + 章节识别 + 表格还原...")
    text = reorganize(text, dedup=dedup, preserve_blanks=True)
    text = split_cover_affiliations(text)
    # Step 5
    p("步骤 5/7: 错误标题降级 + 页脚清理 + 去页码 + 修复跑版段落...")
    text = demote_false_headings(text)
    text = drop_lonely_pagefooter_dupes(text)
    text = strip_heading_page_numbers(text)
    text = fix_runon_paragraphs(text)
    text = merge_cross_page_fragments(text)
    text = split_inline_layout_markers(text)
    text = move_page_split_footnotes(text)
    text = split_body_sentences(text)
    text = convert_numbered_subsection_labels(text)
    text = split_event_entries(text)
    text = normalize_toc_blocks(text)
    text = normalize_figure_captions(text)
    # Step 6
    p("步骤 6/7: 编号标题归一 + 附件提取...")
    text = normalize_numbered_headings(text)
    text = extract_attachments(text)
    # Step 7
    p("步骤 7/7: 页面标记保留 + 规范化空行...")
    text = preserve_page_markers(text)
    text = merge_cross_page_fragments(text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = normalize_blank_lines(text)
    return text.strip() + '\n'


def reorganize(md_text, dedup=True, preserve_blanks=True):
    lines = md_text.split('\n')
    output_lines = []
    headings_seen = set()

    # [Bug③修复 20260926] 首部空行导致崩溃：prev_blank 原本只在循环体内首次赋值，
    # 而 `^---+$`（分隔线）与 RE_DUP_HEAD（"## 第X页"）两个分支会 continue 且不赋值，
    # 所以只要「空行出现在 prev_blank 首次赋值之前」（例如 md 以空行开头）就会抛
    # UnboundLocalError: cannot access local variable 'prev_blank'。
    # 初始置 True（语义＝"行首已在空行之后"）→ 丢弃首部空行，符合 preserve_blanks 原意。
    prev_blank = True

    for line in lines:
        stripped = line.strip()

        if re.match(r'^---+$', stripped):
            continue

        if RE_DUP_HEAD.match(stripped):
            continue

        if not stripped:
            if preserve_blanks and not prev_blank:
                output_lines.append('')
            prev_blank = True
            continue

        cleaned = clean_line(stripped)
        if not cleaned:
            prev_blank = False
            continue

        level, heading_text = classify_section(cleaned)
        prev_blank = False

        if level > 0:
            if dedup:
                unique, headings_seen = deduplicate_headings(headings_seen, level, heading_text)
                if not unique:
                    continue

            prefix = '#' * (level + 1)
            output_lines.append(f"{prefix} {heading_text}")
            output_lines.append('')
        else:
            output_lines.append(cleaned)

    return '\n'.join(output_lines)

# ========== 附件专项整理 ==========

def extract_attachments(md_text):
    lines = md_text.split('\n')
    output_lines = []
    in_attachment = False
    attach_num = 0
    attach_lines = []
    table_buf = []

    def flush_table(buf, out):
        if not buf: return []
        rows = [re.split(r'\s{2,}', r) for r in buf if r.strip()]
        if len(rows) >= 2:
            md_t = try_build_md_table(rows)
            if md_t:
                out.extend(md_t); out.append('')
            else:
                out.extend(buf); out.append('')
        else:
            out.extend(buf); out.append('')
        return []

    for line in lines:
        stripped = line.strip()

        m = RE_ATTACH.match(stripped)

        # [Bug②修复 20260906] 附件模式退出：遇到 H1/H2 标题时结束附件模式
        # 原因：原版 in_attachment 永远不重置，导致后续所有空行被吞、表格被粘连
        if in_attachment and re.match(r'^#{1,2}\s', stripped) and not m:
            if attach_lines:
                table_buf = flush_table(table_buf, attach_lines)
                output_lines.extend(attach_lines)
                output_lines.append('')
                attach_lines = []
            table_buf = []
            in_attachment = False
            output_lines.append(line)
            continue

        if m:
            if attach_lines:
                table_buf = flush_table(table_buf, attach_lines)
                output_lines.extend(attach_lines)
                output_lines.append('')
                attach_lines = []

            in_attachment = True
            attach_num += 1
            attach_lines.append(f"### 附件 {attach_num}")
            rest = stripped[m.end():].strip()
            if rest:
                attach_lines.append(rest)
            continue

        if not in_attachment:
            output_lines.append(line)
            continue

        parts = re.split(r'\s{2,}', stripped) if stripped else []
        is_tbl = (len(parts) >= 2 and max(len(p) for p in parts) < 20 if parts else False)

        if is_tbl:
            table_buf.append(stripped)
        else:
            table_buf = flush_table(table_buf, attach_lines)
            if stripped:
                attach_lines.append(stripped)

    if attach_lines:
        table_buf = flush_table(table_buf, attach_lines)
        output_lines.extend(attach_lines)

    return '\n'.join(output_lines)

# ========== 项目 Markdown 规范（球阀生产管理部）==================

def format_project_md(md_text):
    """套用项目 Markdown 规范：
    1) 去除每行行首的 '>' 块引用标记（仅去标记，正文中的 '>' 如箭头不动）；
    2) 将 '- ' 项目符号列表（含缩进子项）转为 '1）2）3）4）' 全角编号；
    3) 同一列表内部的空行（上下均为列表项的空行）直接删除，避免编号因空行“重头”。
    注意：遇到非空非列表行（标题、正文段等）编号重新从 1 开始。
    """
    raw_lines = md_text.split('\n')
    # 1) 先去掉行首 > 块引用标记
    stripped = [re.sub(r'^\s*>\s*', '', ln) for ln in raw_lines]
    # 2) 合并“夹在两个列表项之间的空行”，保持编号连续
    merged = []
    n = len(stripped)
    for i, ln in enumerate(stripped):
        if ln.strip() == '':
            prev_bullet = i > 0 and re.match(r'^\s*-\s+', stripped[i - 1])
            next_bullet = i < n - 1 and re.match(r'^\s*-\s+', stripped[i + 1])
            if prev_bullet and next_bullet:
                continue  # 删除该空行，避免编号重头
        merged.append(ln)
    # 3) 连续编号：遇到非空非列表行则重置
    out = []
    counter = 0
    prev_bullet = False
    for line in merged:
        if re.match(r'^\s*-\s+', line):
            if not prev_bullet:
                counter = 1
            else:
                counter += 1
            text_body = re.sub(r'^\s*-\s+', '', line)
            out.append(str(counter) + '）' + text_body)
            prev_bullet = True
        else:
            out.append(line)
            prev_bullet = False
    return '\n'.join(out)


# ========== 主程序 ==========

# ========== 清理 PDF 页眉页脚残留 ==========

def _fix_cover_fragments(md_text):
    """修复封面残片：'# 标题（' + 作者 + '# ）' 等提取阶段产生的括号碎片。

    - '# XXX（' / '# ）' 形式 -> 去掉 '#' 与孤立括号，保留为普通行（由后续合并成标题）
    - 单独成行的 '（' / '）' 直接删除
    """
    lines = md_text.split('\n')
    out = []
    for line in lines:
        s = line.strip()
        if re.fullmatch(r'#\s*[（(]', s) or re.fullmatch(r'#\s*[）)]', s):
            continue
        m = re.match(r'^(#{1,3})\s*([（(])(.+)$', s) or re.match(r'^(#{1,3})\s*(.+)[（(]$', s)
        if m:
            # '# 中国绿色燃料发展报告（' -> 把括号并入标题前内容（保留正文）
            prefix = m.group(1)
            rest = re.sub(r'^#{1,3}\s*', '', s)
            out.append(prefix + ' ' + rest.rstrip('（(').strip())
            continue
        out.append(line)
    return '\n'.join(out)


def strip_pdf_headers_footers(md_text):
    """清理文本型 PDF 常见的页眉/页脚残留：
      - 行首 '08|中国绿色燃料发展报告（2026）正文…' -> 去掉 'N|标题' 前缀（保留正文）
      - 整行 '26|中国绿色燃料发展报告（2026）' -> 删除
      - 行 '章标题|页码'（如 '国内外发展情况|07'、'2026年中国绿色燃料发展展望|23'）-> 去掉 '|页码'
      - 孤立页码行（如 '02' / '25'）保留（可能在正文中意义不大，但保守处理）
    仅处理形如 'NN|文字' 或 '文字|NN' 的清晰残留，避免误伤正文中的竖线。
    """
    lines = md_text.split('\n')
    out = []
    header_prefix = re.compile(r'^(\d{1,3})\|(中国绿色燃料发展报告（2026）)(.*)$')
    page_suffix  = re.compile(r'^(.*?)\|(\d{1,3})[①-⑩]*$')
    for line in lines:
        s = line.strip()
        if not s:
            out.append(line)
            continue
        # 行首页码|标题：剥掉页码前缀（保留其余）
        m = header_prefix.match(s)
        if m:
            rest = m.group(3).strip()
            if rest:
                out.append(rest)
            continue
        # 整行 = 页码|标题（无正文）
        if re.fullmatch(r'\d{1,3}\|中国绿色燃料发展报告（2026）', s):
            continue
        # 独立页眉行（尤其中间页/末页残留）
        if re.fullmatch(r'中国绿色燃料发展报告（2026）', s):
            continue
        # 章标题|页码
        m2 = page_suffix.match(s)
        if m2 and 1 <= int(m2.group(2)) <= 2000:
            head = m2.group(1).strip()
            # 只有头部是中文标题（不含句号/很长正文）才去掉页码
            if head and len(head) <= 30 and not re.search(r'[。！？；]', head):
                out.append(head)
                continue
        out.append(line)
    return '\n'.join(out)


# ========== 合并断裂标题 ==========

def merge_broken_headings(md_text):
    """合并断裂的多行标题。
    例如：# 一 / # 国内外发展 / # 情况 → # 一、国内外发展情况
    规则：连续多行都是 # 开头、每行都短（≤40字）、首行无标点。
    """
    lines = md_text.split('\n')
    out = []
    buf = []

    def flush_buf():
        nonlocal buf
        if not buf:
            return
        if len(buf) == 1:
            out.append(buf[0])
        else:
            first = buf[0]
            m = re.match(r'^(#{1,6})\s+', first)
            prefix = m.group(1) if m else '#'
            texts = []
            for b in buf:
                t = re.sub(r'^#{1,6}\s+', '', b).strip()
                if t:
                    texts.append(t)
            # 中文序号（一、二、…）后补顿号；末尾孤立中文数字（页码残留）不拼入
            texts_clean = []
            for idx, t in enumerate(texts):
                if idx == len(texts) - 1 and re.fullmatch(r'[一二三四五六七八九十百]+', t):
                    continue
                texts_clean.append(t)
            if not texts_clean:
                texts_clean = texts
            merged_text = ''
            for idx, t in enumerate(texts_clean):
                if idx == 0 and re.fullmatch(r'[一二三四五六七八九十百]+', t):
                    merged_text += t + '、'
                else:
                    merged_text += t
            merged = prefix + ' ' + merged_text
            out.append(merged)
        buf = []

    for line in lines:
        s = line.strip()
        if not s:
            flush_buf()
            out.append(line)
            continue
        # 检查是否是可合并的标题片段
        m = re.match(r'^#{1,3}\s+(.{1,40})$', s)
        if m:
            inner = m.group(1).strip()
            # 不含句末标点且不是完整标题结构 → 可能是片段
            if not re.search(r'[。！？；]$', inner) and not re.match(r'^[（(][一二三四五六七八九十]+[）)]', inner):
                buf.append(line)
                continue
        flush_buf()
        out.append(line)
    flush_buf()
    return '\n'.join(out)


# ========== 拆分标题正文粘连 ==========

def split_heading_body(md_text):
    """拆分标题与正文粘连的行。

    对（X）开头的标题（如 '（一）全球绿色燃料发展情况国际海事组织…'、
    '（一）坚持系统谋划，加强顶层设计和规划布局制定出台…'），
    以标题结尾词（情况/现状/发展/体系/政策/建设/应用/创新/模式/布局/
    环境/市场/保障/目标/任务/要求/机制/路径/标准/能力/竞争力/篇/章）为拆分点。

    规则：
      - 仅处理以（X）开头的标题行（避免误拆正文行，如 '### 国际民航组织…'）；
      - 强分节词（篇/章）后必为正文，直接拆分；
      - 标题通常在前 45 字内；
      - 若标题词后紧跟正文开头动词（针对/制定/继续/重点/支持/推动/加快/
        组织/开展/聚焦/鼓励/加大/实施/探索/统筹/坚持/立足），在动词前拆分；
      - 否则取前 45 字内最后一个标题词。
    """
    lines = md_text.split('\n')
    out = []
    for line in lines:
        s = line.strip()
        if not s or not s.startswith('#'):
            out.append(line)
            continue
        content = re.sub(r'^#{1,6}\s+', '', s)
        if len(content) <= 50:
            out.append(line)
            continue
        if not re.match(r'^[（(][一二三四五六七八九十]+[）)]', content):
            out.append(line)
            continue

        prefix = '#' * (len(s) - len(s.lstrip('#')))
        split_at = None

        title_words = (r'情况|现状|展望|发展|体系|政策|建设|应用|创新|模式|'
                       r'布局|环境|市场|保障|目标|任务|要求|机制|路径|标准|'
                       r'能力|探索|推进|竞争力|两个市场|篇|章')
        body_start = (r'针对|制定|继续|重点|支持|推动|加快|组织|开展|聚焦|'
                      r'鼓励|加大|实施|探索|统筹|坚持|立足|完善|建立|构建|'
                      r'促进|引导|出台|重大|系列|有关|要|将|加快|研究')

        mP = re.match(r'^[（(][一二三四五六七八九十]+[）)]', content)
        title_start = mP.end()
        rest = content[title_start:]
        # 强分节：篇/章 后直接是正文
        mq = re.search(r'(篇|章)', rest)
        if mq and len(rest) - (mq.end()) >= 3:
            split_at = title_start + mq.end()
        else:
            # 在前 45 字内找标题词
            head = rest[:45]
            matches = list(re.finditer(title_words, head))
            if matches:
                # 优先：标题词后紧跟正文开头动词 → 在动词前拆分
                best = None
                for m in matches:
                    after = rest[m.end():]
                    mv = re.match(body_start, after)
                    if mv:
                        between = rest[m.end():m.end() + mv.end()]
                        if not re.search(r'[，。；！？、]', between):
                            best = title_start + m.end()
                            break
                if best is not None:
                    split_at = best
                else:
                    split_at = title_start + matches[-1].end()

        if split_at and split_at > 3 and split_at < len(content):
            heading_part = content[:split_at]
            body_part = content[split_at:].lstrip(' 　')
            out.append(prefix + ' ' + heading_part)
            out.append('')
            out.append(body_part)
        else:
            out.append(line)
    return '\n'.join(out)


# ========== 正文长行内（X）标题粘连拆分 ==========

def split_runon_body(md_text):
    """正文长行内若出现独立的（X）标题（如 '…高质量发展。（一）国内外发展情况国际海事组织…'），
    将其拆分：标题 = '（X）…标题词'，正文 = 标题词后的部分。

    用于一步提取/转换阶段把「正文 + 标题 + 正文」挤在同一行的情况。
    对同一行内多个（X）标题循环拆分（如 '…公平竞争。（三）…竞争力针对…（四）…'）。
    """
    lines = md_text.split('\n')
    out = []
    title_words = (
        r'情况|现状|展望|发展|体系|政策|建设|应用|创新|模式|'
        r'布局|环境|市场|保障|目标|任务|要求|机制|路径|标准|'
        r'能力|探索|推进|竞争力|两个市场|篇|章'
    )
    title_re = re.compile(title_words)
    body_start_verb = (r'针对|制定|继续|重点|支持|推动|加快|组织|开展|聚焦|'
                       r'鼓励|加大|实施|探索|统筹|坚持|立足|完善|建立|构建|'
                       r'促进|引导|出台|重大|系列|有关|要|将|加快|研究')
    verb_re = re.compile(body_start_verb)
    bracket_re = re.compile(r'[（(][一二三四五六七八九十]+[）)]')

    def find_split(s, from_idx):
        """从 from_idx 开始找第一个（X）标题，返回 (pos, split_pos) 绝对位置或 None。"""
        for m in bracket_re.finditer(s, from_idx):
            pos = m.start()
            if m.end() >= len(s):
                continue
            prev_char = s[pos - 1] if pos > 0 else ''
            if prev_char and prev_char not in '。；，、！？':
                continue
            title_end = m.end()
            rest = s[title_end:]
            if len(rest) < 3:
                continue
            m2s = list(title_re.finditer(rest[:45]))
            if not m2s:
                continue
            split_pos = None
            # 强分节：篇/章 后直接是正文 → 最短标题（如 '技术创新篇个项目…' -> '技术创新篇'）
            mz = re.search(r'(篇|章)', rest[:45])
            if mz and len(rest) - mz.end() >= 3:
                split_pos = title_end + mz.end()
            else:
                for cand in m2s:
                    after = rest[cand.end():]
                    if not after:
                        continue
                    mv = verb_re.match(after)
                    if mv:
                        between = rest[cand.end():cand.end() + mv.end()]
                        if not re.search(r'[，。；！？、]', between):
                            split_pos = title_end + cand.end()
                            break
            if split_pos is None:
                split_pos = title_end + m2s[-1].end()
            return pos, split_pos
        return None

    for line in lines:
        s = line.strip()
        if not s or s.startswith('#') or len(s) <= 50:
            out.append(line)
            continue
        segments = []
        cursor = 0
        first = True
        while cursor < len(s):
            r = find_split(s, cursor)
            if not r:
                break
            pos, split_pos = r
            pre = s[cursor:pos].strip()
            heading = s[pos:split_pos].strip()
            if pre and (first or pre):
                segments.append(pre)
            segments.append('### ' + heading)
            segments.append('')
            cursor = split_pos
            first = False
        if cursor < len(s):
            tail = s[cursor:].strip()
            if tail:
                segments.append(tail)
        if not segments:
            out.append(line)
        else:
            out.extend(segments)
    return '\n'.join(out)


# ========== 错误标题降级（提取阶段误标） ==========

def demote_false_headings(md_text):
    """把提取阶段误标为标题的正文行降级为普通段落。

    条件（须同时满足，避免误伤真标题）：
      - 行以 ### 或 #### 开头（## 及以上太像是真章节，不动）；
      - 内容不含（X）中文序号、不含 X.Y 数字编号、不以 第X章/节 开头；
      - 内容含中文逗号（，）或句号（。），说明是正文句子而非标题；
      - 内容长度 > 12 且以中文字符开头。
    """
    lines = md_text.split('\n')
    out = []
    for line in lines:
        s = line.strip()
        if not s.startswith('###'):
            out.append(line)
            continue
        level = len(s) - len(s.lstrip('#'))
        if level < 3:
            out.append(line)
            continue
        content = re.sub(r'^#{3,}\s+', '', s)
        if (re.match(r'^[（(][一二三四五六七八九十]+[）)]', content)
                or re.match(r'^\d+(?:\.\d+)*', content)
                or re.match(r'^第[一二三四五六七八九十百千]+[章节篇部]', content)
                or not re.match(r'^[一-鿿]', content)
                or len(content) <= 12):
            out.append(line)
            continue
        if re.search(r'[，。；！？]', content):
            out.append(content)
            continue
        out.append(line)
    return '\n'.join(out)


# ========== 删除孤立页脚重复行 ==========

def drop_lonely_pagefooter_dupes(md_text):
    """删除页脚残留的孤立行：独立成行、前后（或相邻）为空行、
    且其文本与文档中 # 标题行重复（如 '结束语'、'大事记'、'国内外发展情况'）。"""
    def norm(t):
        # 去标题前缀：'一、' / '（一）' / '1.' 等
        t = re.sub(r'^[（(][一二三四五六七八九十]+[）)]\s*', '', t)
        t = re.sub(r'^[一二三四五六七八九十百]+[、.．]?\s*', '', t)
        t = re.sub(r'^\d+(?:\.\d+)*[.．、]?\s*', '', t)
        return t.strip()

    lines = md_text.split('\n')
    # 收集标题文本（含规范化形式）
    titles = set()
    for ln in lines:
        s = ln.strip()
        if s.startswith('#'):
            raw_t = re.sub(r'^#{1,6}\s+', '', s).strip()
            titles.add(raw_t)
            titles.add(norm(raw_t))
    # 查找孤立行（非标题、前后为空或页标记的短行）
    n = len(lines)
    drop = set()
    for i in range(n):
        s = lines[i].strip()
        if not s or s.startswith('#') or s.startswith('![') or s.startswith('<!--'):
            continue
        if s not in titles and norm(s) not in titles:
            continue
        prev_ok = i == 0 or not lines[i-1].strip()
        next_ok = i == n-1 or not lines[i+1].strip()
        # 允许前一行是 <!-- 页标记 -->
        if i > 0 and lines[i-1].strip().startswith('<!--'):
            prev_ok = True
        if i < n-1 and lines[i+1].strip().startswith('<!--'):
            next_ok = True
        if prev_ok and next_ok:
            drop.add(i)
    out = [ln for i, ln in enumerate(lines) if i not in drop]
    return '\n'.join(out)


# ========== 加粗子标题转标题 ==========

def convert_bold_subheadings(md_text):
    """将 **① xxx** 和 **② xxx** 等加粗子标题转为 ### 标题。"""
    lines = md_text.split('\n')
    out = []
    bold_re = re.compile(r'^\*\*([①②③④⑤⑥⑦⑧⑨⑩]\s*.+?)\*\*\s*$')
    for line in lines:
        s = line.strip()
        m = bold_re.match(s)
        if m:
            text = m.group(1).strip()
            if len(text) >= 3:
                out.append('### ' + text)
            else:
                out.append(line)
        else:
            out.append(line)
    return '\n'.join(out)


# ========== 去除标题末尾页码 ==========

def strip_heading_page_numbers(md_text):
    """去除标题末尾的页码数字（如 22、23）和 |N 标记。"""
    lines = md_text.split('\n')
    out = []
    for line in lines:
        s = line.strip()
        if not s.startswith('#'):
            out.append(line)
            continue
        content = re.sub(r'^#{1,6}\s+', '', s)
        new_content = re.sub(r'\|\s*\d{1,3}[①-⑩]*\s*$', '', content).strip()
        new_content = re.sub(r'\s+\d{1,3}\s*$', '', new_content).strip()
        new_content = re.sub(r'(?<=[一-鿿）)])\d{1,3}\s*$', '', new_content).strip()
        if new_content != content:
            prefix = '#' * (len(s) - len(s.lstrip('#')))
            out.append(prefix + ' ' + new_content)
        else:
            out.append(line)
    return '\n'.join(out)


# ========== 修复跑版段落 ==========

def fix_runon_paragraphs(md_text):
    """合并正文中未结束句子的断行，保留结构性边界。"""
    lines = md_text.split('\n')
    out = []
    cover_units = {
        '国家能源局能源节约和科技装备司',
        '中国产业发展促进会',
        '电力规划设计总院',
        '石油和化学工业规划院',
    }
    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            out.append(lines[i])
            i += 1
            continue
        if s.startswith(('#', '!', '<!--', '◎')) or s in cover_units:
            out.append(lines[i])
            i += 1
            continue
        if i + 1 < len(lines):
            next_s = lines[i + 1].strip()
            protected_next = (
                not next_s
                or next_s.startswith(('#', '!', '<!--', '|', '>', '```', '◎'))
                or next_s in cover_units
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]+(?:\s|$)', next_s)
                or re.match(r'^图(?:\d+|我国|全球)', next_s)
            )
            protected_current = (
                re.match(r'^图(?:\d+|我国|全球)', s)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]+(?:\s|$)', s)
            )
            if (next_s and not protected_next and not protected_current
                    and not re.search(r'[。！？；：]$', s)):
                out.append(lines[i].rstrip() + next_s)
                i += 2
                continue
        out.append(lines[i])
        i += 1
    return '\n'.join(out)


def merge_cross_page_fragments(md_text):
    """合并被页标、页首图片截断的同一句文本。"""
    lines = md_text.split('\n')
    out = []
    cover_units = {
        '国家能源局能源节约和科技装备司',
        '中国产业发展促进会',
        '电力规划设计总院',
        '石油和化学工业规划院',
    }
    page_re = re.compile(r'^<!--\s*第\s*\d+\s*页\s*-->$')
    i = 0
    while i < len(lines):
        current = lines[i].strip()
        if (current and current not in cover_units
                and not current.startswith(('#', '!', '<!--', '◎'))
                and not re.search(r'[。！？；：]$', current)):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and page_re.match(lines[j].strip()):
                k = j + 1
                while k < len(lines):
                    value = lines[k].strip()
                    if value and not value.startswith(('!', '<!--')):
                        break
                    k += 1
                if k < len(lines):
                    continuation = lines[k].strip()
                    continuation = re.sub(
                        r'^[①②③④⑤⑥⑦⑧⑨⑩](?=(?:关部门|再生能源|色航空燃料))',
                        '', continuation,
                    )
                    if (continuation and not continuation.startswith(('#', '◎'))
                            and not re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]+(?:\s|$)', continuation)):
                        out.extend(lines[i + 1:k])
                        out.append(current + continuation)
                        i = k + 1
                        continue
        out.append(lines[i])
        i += 1
    return '\n'.join(out)


def split_cover_affiliations(md_text):
    """拆分封面署名串，避免多个编著单位被 OCR 串成一行。"""
    affiliation = '国家能源局能源节约和科技装备司中国产业发展促进会电力规划设计总院石油和化学工业规划院'
    units = [
        '国家能源局能源节约和科技装备司',
        '中国产业发展促进会',
        '电力规划设计总院',
        '石油和化学工业规划院',
    ]
    lines = md_text.split('\n')
    out = []
    for line in lines:
        content = re.sub(r'^#{1,6}\s+', '', line.strip())
        if content == affiliation:
            out.extend(units)
        else:
            out.append(line)
    return '\n'.join(out)


def split_inline_layout_markers(md_text):
    """分离正文中的图题、脚注说明和孤立脚注角标。"""
    lines = []
    for line in md_text.split('\n'):
        if re.fullmatch(r'\s*[①②③④⑤⑥⑦⑧⑨⑩]+\s*', line):
            continue
        value = line.strip()
        if value:
            value = re.sub(r'(?<!^)(?=图(?:\d+|我国|全球))', '\n', value)
            caption = re.compile(
                r'图(?:\d+|我国|全球)[^。！？；]*?(?:情况|产量（[^）]+）)'
            )
            value = caption.sub(lambda match: match.group(0) + '\n', value)
            footnote_leads = (
                r'的名字源自|绿色甲醇主要以|绿色合成氨主要以|绿色航空燃料主要以|'
                r'生物燃料乙醇主要以|生物柴油主要以|生物天然气主要以|'
                r'结合|引自|中国产业发展促进会|石油和化学工业规划院|'
                r'共生产模式|质量平衡法'
            )
            value = re.sub(
                r'(?<![①②③④⑤⑥⑦⑧⑨⑩])([①②③④⑤⑥⑦⑧⑨⑩])(?=(?:' + footnote_leads + r'))',
                r'\n\1 ',
                value,
            )
            value = re.sub(r'([一-鿿])[ \t　]+[①②③④⑤⑥⑦⑧⑨⑩][ \t　]+([一-鿿])', r'\1\2', value)
            value = re.sub(r'(?<=[一-鿿0-9])[①②③④⑤⑥⑦⑧⑨⑩](?=[一-鿿])', '', value)
            value = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩](?=[关再色])', '', value)
            value = re.sub(r'\n+', '\n', value).strip()
            lines.extend(value.split('\n'))
        else:
            lines.append(line)
    return re.sub(r'\n{3,}', '\n\n', '\n'.join(lines))


def move_page_split_footnotes(md_text):
    """把页尾脚注移到下一页续句之后，避免脚注插入正文中间。"""
    lines = md_text.split('\n')
    out = []
    i = 0
    page_re = re.compile(r'^<!--\s*第\s*\d+\s*页\s*-->$')
    footnote_re = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]\s+.+')
    while i < len(lines):
        current = lines[i].strip()
        if current and not current.startswith(('#', '!', '<!--')) and not re.search(r'[。！？；：]$', current):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and footnote_re.match(lines[j].strip()):
                footnote = lines[j].strip()
                k = j + 1
                while k < len(lines) and not lines[k].strip():
                    k += 1
                if k < len(lines) and page_re.match(lines[k].strip()):
                    m = k + 1
                    while m < len(lines) and not lines[m].strip():
                        m += 1
                    if m < len(lines) and not lines[m].strip().startswith(('#', '!', '<!--')):
                        out.extend([lines[i].rstrip(), '', lines[k].strip(), '', lines[m].strip(), '', footnote])
                        i = m + 1
                        continue
        out.append(lines[i])
        i += 1
    return '\n'.join(out)


def split_body_sentences(md_text):
    """按句末标点拆分长正文，避免整页内容挤成单行。"""
    out = []
    for line in md_text.split('\n'):
        stripped = line.strip()
        if not stripped or stripped.startswith(('#', '!', '<!--', '|', '>', '```', '◎')):
            out.append(line)
            continue
        parts = re.split(r'(?<=[。！？；])(?=[^」』）)\]])', stripped)
        if len(parts) > 1:
            out.extend(part.strip() for part in parts if part.strip())
        else:
            out.append(line)
    return '\n'.join(out)


FUEL_SUBSECTION_LABELS = '绿色甲醇|绿色合成氨|绿色航空燃料|生物燃料乙醇|生物柴油|生物天然气'


def convert_numbered_subsection_labels(md_text):
    """将绿色燃料品种的角标小标题/独立残片统一为三级标题。"""
    labels = FUEL_SUBSECTION_LABELS
    pattern = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]+\s*(' + labels + r')。(.*)$')
    exact_pattern = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]+\s*(' + labels + r')。?$')
    plain_pattern = re.compile(r'^(' + labels + r')。$')
    out = []
    for line in md_text.split('\n'):
        stripped = line.strip()
        match = pattern.match(stripped)
        if match:
            out.append(f'### {match.group(1)}')
            if match.group(2).strip():
                out.extend(['', match.group(2).strip()])
        elif exact_pattern.match(stripped):
            match = exact_pattern.match(stripped)
            out.append(f'### {match.group(1)}')
        elif plain_pattern.match(stripped):
            out.append(f'### {plain_pattern.match(stripped).group(1)}')
        else:
            out.append(line)
    return '\n'.join(out)


def split_event_entries(md_text):
    """拆分“大事记”中标题与日期正文粘连的条目。"""
    lines = md_text.split('\n')
    out = []
    in_events = False
    date_re = re.compile(r'20\d{2}年(?:\d{1,2}月)?\s*[，,]')
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('# 大事记'):
            in_events = True
            out.append(line)
            continue
        if in_events and stripped.startswith('#') and not stripped.startswith('###'):
            in_events = False
        if not in_events or not stripped or stripped.startswith(('#', '!', '<!--')):
            out.append(line)
            continue

        date_match = date_re.search(stripped)
        if date_match and date_match.start() > 0:
            title = stripped[:date_match.start()].strip()
            body = stripped[date_match.start():].strip()
            out.extend([f'**{title}**', body])
            continue

        marker = '国家能源局公布'
        if marker in stripped and not stripped.startswith(marker):
            title, body = stripped.split(marker, 1)
            out.extend([f'**{title.strip()}**', marker + body])
            continue
        out.append(line)
    return '\n'.join(out)


def normalize_toc_blocks(md_text):
    """将目录页 OCR 串行文本还原为普通嵌套列表。"""
    lines = md_text.split('\n')
    out = []
    toc_seen = False
    for line in lines:
        stripped = line.strip()
        if stripped == '# 目录':
            toc_seen = True
            out.append(line)
            continue
        if toc_seen and stripped.startswith('# 一、国内外发展情况'):
            toc_seen = False
        if stripped.startswith('一国内外发展情况'):
            out.extend([
                '- 一、国内外发展情况',
                '  - （一）全球绿色燃料发展情况',
                '    - 1. 燃料转型加速推进',
                '    - 2. 总体产能稳中有升',
                '    - 3. 产业创新持续深化',
                '  - （二）国内绿色燃料发展情况',
                '    - 1. 政策框架逐步建立',
                '    - 2. 产业发展初具规模',
                '    - 3. 应用场景逐步拓展',
                '    - 4. 创新体系不断健全',
            ])
            continue
        if toc_seen and stripped.startswith('二2026年中国绿色燃料发展展望'):
            out.extend([
                '- 二、2026年中国绿色燃料发展展望',
                '  - （一）坚持系统谋划，加强顶层设计和规划布局（第22页）',
                '  - （二）坚持试点先行，探索形成产业科学发展模式（第23页）',
                '  - （三）坚持创新引领，不断提升产业发展竞争力（第23页）',
                '  - （四）坚持优化环境，有力支撑产业规模化发展（第23页）',
                '  - （五）坚持需求牵引，统筹国际国内两个市场（第24页）',
                '- 结束语',
                '- 大事记',
                '  - （一）政策规范篇',
                '  - （二）技术创新篇',
                '  - （三）工程建设篇',
            ])
            continue
        out.append(line)
    return '\n'.join(out)


def normalize_figure_captions(md_text):
    """补齐图题编号与图号后的空格，修复 OCR 丢号/粘连/图号拆分。"""
    replacements = {
        '图我国绿色甲醇产能分布情况': '图7 我国绿色甲醇产能分布情况',
        '图我国绿色航空燃料产能分布情况': '图9 我国绿色航空燃料产能分布情况',
        '图我国生物柴油产能分布情况': '图11 我国生物柴油产能分布情况',
        '图1 1 我国生物柴油产能分布情况': '图11 我国生物柴油产能分布情况',
    }
    labels = FUEL_SUBSECTION_LABELS
    out = []
    for line in md_text.split('\n'):
        value = line.strip()
        if value == 'ICAO' or re.fullmatch(r'[①②③④⑤⑥⑦⑧⑨⑩]+', value):
            continue
        value = replacements.get(value, value)
        # 图题后粘连下一品种小标题：`图10 …生物柴油。` → 图题 + 空行 + `### 生物柴油`
        split_m = re.match(r'^(图\d+\s+[^。]+)(' + labels + r')。$', value)
        if split_m:
            out.extend([split_m.group(1), '', f'### {split_m.group(2)}'])
            continue
        # 图号与四位数年份粘连：`图32020 年以来…` → `图3 2020 年以来…`
        value = re.sub(r'^图(\d{1,2})\s*(\d{4})\s*年', r'图\1 \2 年', value)
        # 图号内数字被拆开：`图1 1 …` → `图11 …`
        value = re.sub(r'^图(\d)\s+(\d)\s*(?=[一-鿿])', r'图\1\2 ', value)
        # 图号后统一补空格
        value = re.sub(r'^(图\d+)(?=[一-鿿])', r'\1 ', value)
        out.append(value if value != line.strip() else line)
    return '\n'.join(out)


# ========== 页面标记保留 ==========

def preserve_page_markers(md_text):
    """保留<!-- 第 X 页 -->并确保前后有空行。"""
    lines = md_text.split('\n')
    out = []
    page_re = re.compile(r'^<!--\s*第\s*\d+\s*页\s*-->$')
    for i, line in enumerate(lines):
        s = line.strip()
        if page_re.match(s):
            if out and out[-1].strip():
                out.append('')
            out.append(line)
            out.append('')
        else:
            out.append(line)
    return '\n'.join(out)


# ========== 规范化空行 ==========

def normalize_blank_lines(md_text):
    """规范化空行：段落间保留一个空行，标题/图片前后各一个空行。"""
    lines = md_text.split('\n')
    out = []
    prev_blank = False
    for line in lines:
        s = line.strip()
        is_blank = not s
        if s.startswith('#') and out and out[-1].strip():
            out.append('')
        if s.startswith('![') and out and out[-1].strip():
            out.append('')
        if is_blank:
            if not prev_blank:
                out.append(line)
        else:
            out.append(line)
        prev_blank = is_blank
    return '\n'.join(out)



def main():
    ap = argparse.ArgumentParser(description="Markdown 整理（PDF/OCR 后处理）")
    ap.add_argument("input", help="输入 Markdown 文件（原始 PDF 转出稿）")
    ap.add_argument("--output", "-o", default=None,
                    help="输出 Markdown 路径（默认**原地更新**输入文件）")
    ap.add_argument("--no-dedup", action="store_true", help="关闭章节去重")
    ap.add_argument("--proj-md", action="store_true",
                    help="套用项目 Markdown 规范：去行首 > 块引用、'- '列表转 1）2）3）4）全角编号、列表内空行合并防重头")
    args = ap.parse_args()

    in_path = os.path.abspath(args.input)
    if not os.path.exists(in_path):
        print(f"[错误] 文件不存在: {in_path}")
        return 1

    if args.output:
        out_path = os.path.abspath(args.output)
    else:
        # 默认**原地更新**：整理后的产物就是与源文件同名的那个 .md，
        # 不再另存 `<基名>_整理.md`（避免同一份内容出现两个文件名）。
        out_path = in_path

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    print(f"[Step 2] 读取: {in_path}", flush=True)
    with open(in_path, "r", encoding="utf-8") as f:
        text = f.read()

    original_text = text
    orig_len = len(text)
    print(f"  字符数: {orig_len:,}", flush=True)

    # 全部整理步骤统一走 full_reorganize()，与 convert.py 单文件路径共用同一实现
    text = full_reorganize(text, dedup=not args.no_dedup, verbose=True)

    # Step 2e: 套用项目 Markdown 规范（可选开关 --proj-md）
    if args.proj_md:
        print("  套用项目 Markdown 规范（去行首 > / '- '列表转 1）编号）...", flush=True)
        text = format_project_md(text)

    if out_path == in_path and text == original_text:
        print("\n[完成] 内容无变化，未重写原文件。", flush=True)
    else:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)

    print(f"\n[完成]", flush=True)
    print(f"   输出: {out_path}", flush=True)
    print(f"   字符数: {len(text):,}（整理前 {orig_len:,}）", flush=True)
    return 0

if __name__ == "__main__":
    sys.exit(main())

