# -*- coding: utf-8 -*-
"""文本型 PDF 直取文字层转 Markdown（不渲染图片、不调用大模型）。

所属：omni-md-toolkit（scripts/extract_text_pdf.py）
特性支持：多栏布局（按 x 间隙分栏，再按列聚类）、项目符号（•/Ø/‣/▪/▸/▶/➢ 等）、
子弹点续行合并、编号标题分级、页码/页眉过滤、CID 清理。

适用：pdfplumber 提取 page.chars 占主导的"文本型" PDF。
判断逻辑见 classify()：chars==0 或字符数远小于图片数，判为扫描/图片型，应走 LLM 路线。

标题定级公式与 md_reorganizer.py 统一：编号段数 + 1（段数 = 用 '.' 分隔的段数）。
"""
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

import pdfplumber

# ---- 可调配置 ----
BODY_TOL = 3.0            # 同行 top 容差（pt）
COL_GAP = 47.0            # x 间隙大于此值视为分栏（同一扫描行内两栏的分界）
HEADER_RE = re.compile(r'^\d+(?:\.\d+)*\.?\s')        # 编号标题：1 / 1.1 / 3.4.1
BULLET_RE = re.compile(r'^[n•Ø‣▪▸▶➢]')                  # 项目符号
MARKER_RE = re.compile(r'^n\s*$')                       # 孤立的 n（误识别的 bullet 占位）
TERMINAL = ('。', '！', '？')                           # 句末标点：用于判断子弹点续行是否结束
PAGE_NUM = re.compile(r'^\s*[-–—]?\s*\d{1,3}\s*[-–—]?\s*$')   # 纯页码页脚（1-3 位）
# 常见 CID 映射（按需扩充；pdfplumber 自带 CMap 已能解大部分）
CID_MAP = {
    "(cid:188)": "（", "(cid:189)": "）", "(cid:190)": "—",
}

DECO_CHARS = '—–' + '\\' + '-·　'


def strip_deco(t):
    return t.rstrip(DECO_CHARS + ' \t').rstrip()


def classify(pdf_path):
    """返回 ('text'|'scan', 明细字典)。"""
    with pdfplumber.open(pdf_path) as pdf:
        rows = [(len(p.chars), len(p.images)) for p in pdf.pages]
    n = len(rows)
    total_chars = sum(c for c, _ in rows)
    total_images = sum(im for _, im in rows)
    image_pages = sum(1 for c, im in rows if c == 0 or (im > 0 and c < 20))
    if total_chars == 0:
        return "scan", dict(total_chars=0, total_images=total_images, image_pages=image_pages, n=n)
    if image_pages >= max(1, n // 2):
        return "scan", dict(total_chars=total_chars, total_images=total_images, image_pages=image_pages, n=n)
    if total_images > 0 and total_chars < total_images:
        return "scan", dict(total_chars=total_chars, total_images=total_images, image_pages=image_pages, n=n)
    return "text", dict(total_chars=total_chars, total_images=total_images, image_pages=image_pages, n=n)


def clean(text):
    text = text.replace('　', ' ')
    return re.sub(r'\s+', ' ', text).strip()


def is_decorative(t):
    if re.fullmatch(r'[0-9]{1,4}', t):
        return True
    if len(t) <= 2 and re.fullmatch(r'[A-Za-z0-9]+', t):
        return True
    return False


def clean_cid(text):
    for k, v in CID_MAP.items():
        text = text.replace(k, v)
    return text


def get_lines(page, tol=BODY_TOL, colgap=COL_GAP):
    """按 top 分组、组内按 x 间隙分栏，返回行字典列表与正文字号。"""
    chars = page.chars
    if not chars:
        return [], 12.0
    groups = {}
    for c in chars:
        groups.setdefault(round(c['top'] / tol), []).append(c)
    lines = []
    for key in sorted(groups):
        g = sorted(groups[key], key=lambda c: c['x0'])
        sub = [g[0]]
        for c in g[1:]:
            if c['x0'] - sub[-1]['x1'] > colgap:
                lines.append(sub)
                sub = [c]
            else:
                sub.append(c)
        lines.append(sub)
    out = []
    for sub in lines:
        sub = sorted(sub, key=lambda c: c['x0'])
        text = ''.join(ch['text'] for ch in sub).strip()
        if not text:
            continue
        x0 = min(ch['x0'] for ch in sub)
        top = min(ch['top'] for ch in sub)
        size = max(ch['size'] for ch in sub)
        bold = any(('Bold' in ch['fontname']) or ('Hei' in ch['fontname']) for ch in sub)
        out.append({'x0': x0, 'top': top, 'text': text, 'size': size, 'bold': bold})
    sizes = sorted(c['size'] for c in chars)
    base = sizes[len(sizes) // 2] if sizes else 12.0
    return out, base


def preprocess(lines, base):
    for ln in lines:
        t = clean(ln['text'])
        m = re.match(r'^(\d{1,3})(?=[一-鿿])', t)
        if m:
            t = t[m.end():]
            ln['size'] = base
        ln['ctext'] = t
        ln['marker'] = False
        if not t:
            ln['kind'] = 'skip'
            continue
        if PAGE_NUM.match(t):
            ln['kind'] = 'skip'
            continue
        if MARKER_RE.match(t):
            ln['kind'] = 'skip'
            ln['marker'] = True
            continue
        if is_decorative(t):
            ln['kind'] = 'skip'
            continue
        ln['kind'] = None


def cluster_columns(lines):
    xs = sorted(set(round(ln['x0']) for ln in lines if ln.get('ctext')))
    if not xs:
        return
    clusters = []
    cur = [xs[0]]
    for x in xs[1:]:
        if x - cur[-1] > 70:
            clusters.append(cur)
            cur = [x]
        else:
            cur.append(x)
    clusters.append(cur)
    centers = [sum(c) / len(c) for c in clusters]
    for ln in lines:
        if not ln.get('ctext'):
            continue
        rx = round(ln['x0'])
        best_i = min(range(len(centers)), key=lambda i: abs(centers[i] - rx))
        ln['col'] = best_i


def terminated(buf):
    return buf['text'].rstrip().endswith(TERMINAL)


def flush(buf, out):
    if buf is None:
        return
    if buf['kind'] == 'bullet':
        out.append('- ' + buf['text'])
    else:
        out.append(buf['text'])


def process_column(col_lines, out):
    pending_marker = False
    for ln in col_lines:
        if ln['kind'] == 'skip':
            if ln.get('marker'):
                pending_marker = True
            continue
        t = ln['ctext']
        if pending_marker:
            t2 = BULLET_RE.sub('', t, 1).strip()
            if t2:
                ln['kind'] = 'bullet'
                ln['ctext'] = t2
            else:
                ln['kind'] = 'skip'
            pending_marker = False
            continue
        if BULLET_RE.match(t):
            ln['kind'] = 'bullet'
            ln['ctext'] = BULLET_RE.sub('', t, 1).strip()
            continue
        if HEADER_RE.match(t):
            num = re.split(r'[.．\s]', t.strip())[0]
            seg = len(re.split(r'[.．]', num))
            ln['level'] = min(seg + 1, 6)
            ln['kind'] = 'heading'
            continue
        if ln['bold'] and len(t) <= 30:
            ln['kind'] = 'maybe_heading'
            ln['level'] = 1 if ln['size'] >= 24 else 3
            continue
        ln['kind'] = 'body'
    buf = None
    for ln in col_lines:
        k = ln['kind']
        t = ln['ctext']
        if k == 'skip':
            continue
        if k == 'heading':
            flush(buf, out)
            buf = None
            out.append('#' * ln['level'] + ' ' + strip_deco(t))
        elif k == 'maybe_heading':
            if buf is not None and buf['kind'] in ('body', 'bullet') and not terminated(buf):
                buf['text'] += t
            else:
                flush(buf, out)
                buf = None
                out.append('#' * ln['level'] + ' ' + strip_deco(t))
        elif k == 'bullet':
            flush(buf, out)
            buf = None
            buf = {'kind': 'bullet', 'text': t}
        else:
            if buf is None:
                buf = {'kind': 'body', 'text': t}
            else:
                buf['text'] += t
    flush(buf, out)


def extract_text_to_markdown(pdf_path):
    """直接提取文本型 PDF 返回 Markdown 文本字符串。"""
    md = []
    with pdfplumber.open(pdf_path) as pdf:
        for pi, page in enumerate(pdf.pages, 1):
            lines, base = get_lines(page)
            preprocess(lines, base)
            cluster_columns(lines)
            content = [ln for ln in lines if ln.get('ctext')]
            if not any(ln['kind'] != 'skip' for ln in content):
                continue
            cols = {}
            for ln in content:
                cols.setdefault(ln['col'], []).append(ln)
            page_out = []
            for ci in sorted(cols):
                col_lines = sorted(cols[ci], key=lambda l: (l['top'], l['x0']))
                process_column(col_lines, page_out)
                page_out.append('')
            if page_out:
                md.append('<!-- 第 %d 页 -->' % pi)
                md.extend(page_out)
                md.append('')
    text = '\n'.join(md).strip() + '\n'
    text = re.sub(r'([一-鿿])\.+', r'\1。', text)
    text = re.sub(r'([一-鿿]),([一-鿿])', r'\1，\2', text)
    text = re.sub(r'([一-鿿])\(', r'\1（', text)
    text = re.sub(r'\)([一-鿿])', r'）\1', text)
    text = clean_cid(text)
    return text


def extract(pdf_path, out_path):
    text = extract_text_to_markdown(pdf_path)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(text)
    return len(text.splitlines())


USAGE = """用法: python extract_text_pdf.py <PDF> [输出.md]

文本型 PDF 直取文字层转 Markdown（不渲染图片、不调用大模型）。
  [输出.md] 缺省时，写与源文件同名的 <基名>.md。

选项:
  -h, --help    显示本帮助

注意: 本脚本是**内部引擎**（由 convert.py 按 --pdf-engine extract_text 调用其函数），
     单独运行产出的 md **不会过版式自整理**；要交付请用 convert.py 或 pipeline.py。
"""


def main():
    argv = sys.argv[1:]
    if argv and argv[0] in ("-h", "--help"):
        print(USAGE.rstrip("\n"))
        return 0
    if not argv or argv[0].startswith("-"):
        if argv:
            sys.stderr.write("未知选项: %s\n\n" % argv[0])
        sys.stderr.write(USAGE)
        return 1
    pdf = argv[0]
    kind, info = classify(pdf)
    print("判断: %s | chars=%d images=%d image_pages=%d/%d" % (
        kind, info["total_chars"], info["total_images"], info["image_pages"], info["n"]))
    if kind == "scan":
        print("[跳过] 该 PDF 判为扫描型/图片型，请走 LLM 视觉识别路线（ocr_llm.py）。")
        return 0
    out = argv[1] if len(argv) > 1 else os.path.splitext(pdf)[0] + ".md"
    n = extract(pdf, out)
    print("已提取: %s (%d 行 MD)" % (out, n))
    return 0


if __name__ == "__main__":
    sys.exit(main())
