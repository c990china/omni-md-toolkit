# -*- coding: utf-8 -*-
"""跨格式图片处理：识别 / 切割 / 保存 / 嵌入 Markdown。

子命令：docx / pptx / pdf / html / image
通用选项：
  --out DIR      资源目录（默认 源目录/assets）
  --recognize    生成 alt 文本：**不做 OCR、不调网关**，只列出候选清单并打印作业单，
                 由当前模型看图写 alt（用 --alt-file 落地）
  --alt-file F   按「相对路径<TAB>alt」清单生成嵌入片段（配合 --recognize 的作业单使用）
  --split        按内容空白带自动切割大图
  --emit FILE    把嵌入片段写入该 .md（默认打印到 stdout）
依赖：pillow, pdfplumber(仅pdf), python-docx(仅docx), python-pptx(仅pptx), beautifulsoup4(仅html)
"""
import os, sys, io, argparse, glob
HERE = os.path.dirname(os.path.abspath(__file__))
# 本脚本真实路径（带中文前缀的长文件名）：作业单里必须打印它，短名在磁盘上并不存在。
_SELF_PATH = os.path.abspath(__file__)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    from PIL import Image
except Exception:
    Image = None

ROOT = os.path.dirname(os.path.abspath(__file__))
IMG_EXT = ("png", "jpg", "jpeg", "gif", "bmp", "webp")


def _find_script(name):
    for f in os.listdir(HERE):
        if f.endswith(name + ".py") or f.endswith("_" + name + ".py"):
            return os.path.join(HERE, f)
    return os.path.join(HERE, name + ".py")


def assets_dir(src, out):
    if out:
        d = out
    else:
        base = os.path.abspath(src)
        d = os.path.join(os.path.dirname(base) if os.path.isfile(base) else base, "assets")
    os.makedirs(d, exist_ok=True)
    return d


def save_image(data, assets, stem, idx, ext="png"):
    name = "%s_%02d.%s" % (stem, idx, ext if ext in IMG_EXT else "png")
    path = os.path.join(assets, name)
    with open(path, "wb") as f:
        f.write(data)
    return path


def split_by_whitespace(path):
    """按水平空白带切分；无切分时返回 [原图]。"""
    if Image is None:
        return [path]
    img = Image.open(path).convert("RGB")
    w, h = img.size
    px = img.load()
    white_th = 240
    row_dark = []
    step = max(1, w // 200)
    for y in range(h):
        dark = 0
        for x in range(0, w, step):
            r, g, b = px[x, y]
            if not (r > white_th and g > white_th and b > white_th):
                dark += 1
        row_dark.append(dark)
    bands, in_band = [], False
    for y, d in enumerate(row_dark):
        if d == 0:
            if not in_band:
                in_band, bands = True, bands + [[y, y]]
            else:
                bands[-1][1] = y
        else:
            in_band = False
    cuts = [(b[0] + b[1]) // 2 for b in bands if (b[1] - b[0]) >= 10]
    if not cuts:
        return [path]
    edges = [0] + cuts + [h]
    subs, stem = [], os.path.splitext(path)[0]
    for i in range(len(edges) - 1):
        y0, y1 = edges[i], edges[i + 1]
        if y1 - y0 < 20:
            continue
        sub = img.crop((0, y0, w, y1))
        sp = "%s_part%d.png" % (stem, i)
        sub.save(sp)
        subs.append(sp)
    return subs or [path]


ALT_CANDIDATES_NAME = "_alt_candidates.txt"
ALT_RESULT_NAME = "_alt_result.txt"


def rel_path(path, assets):
    """图片相对 assets 上级目录的引用路径（正斜杠）。"""
    return os.path.relpath(path, os.path.dirname(assets) or ".").replace(os.sep, "/")


def load_alt_map(path):
    """读入「相对路径<TAB>alt」清单。返回 {归一化相对路径: alt}。"""
    alt = {}
    with io.open(path, encoding="utf-8-sig") as f:
        for ln in f:
            ln = ln.rstrip("\n").rstrip("\r")
            if not ln.strip() or ln.lstrip().startswith("#"):
                continue
            parts = ln.split("\t")
            if len(parts) < 2:
                continue
            alt[parts[0].strip().replace("\\", "/")] = parts[1].strip()
    return alt


def write_alt_candidates(assets, paths):
    """把需要 alt 的图片清单写到 assets 目录，返回文件路径。"""
    p = os.path.join(assets, ALT_CANDIDATES_NAME)
    with io.open(p, "w", encoding="utf-8", newline="") as f:
        f.write("# 需要 alt 文本的图片清单（media_extract 生成，仅参考）。\n")
        f.write("# 请把 alt 写成「相对路径<TAB>alt文本」另存为 %s 后运行：--alt-file <文件>\n"
                % ALT_RESULT_NAME)
        for t in paths:
            f.write(rel_path(t, assets) + "\n")
    return p


def alt_handoff_text(assets, paths):
    """生成「请当前模型看图写 alt」的作业单。"""
    result = os.path.join(assets, ALT_RESULT_NAME)
    L = []
    L.append("[图片 alt] 需要为下列 %d 张图片写 alt 文本。" % len(paths))
    L.append("")
    L.append("  ▶ 请由你（当前模型）逐张读取下面的图片，把 alt 写成「相对路径<TAB>alt文本」写入：")
    L.append("       %s" % result)
    L.append("     然后运行：")
    L.append('       python "%s" image "<图片目录>" --alt-file "%s" --emit "<目标.md>"'
             % (_SELF_PATH, result))
    L.append("")
    L.append("  ▶ 本技能**不做 OCR，也不调任何网关**：alt 一律由具备视觉能力的模型生成。")
    L.append("     若你（当前模型）没有视觉能力：请使用者提供一个视觉模型的")
    L.append("     【模型名 / 网关地址 / 密钥】，由你自行调用它生成 alt（本工具不做 OCR、")
    L.append("     不含网关调用代码）；或切换到具备视觉能力的模型后重跑。")
    L.append("     或切换到具备视觉能力的模型后重跑。")
    L.append("")
    L.append("  候选图片：")
    for t in paths[:50]:
        L.append("   - %s" % rel_path(t, assets))
    if len(paths) > 50:
        L.append("   ...（共 %d 张，此处仅列前 50）" % len(paths))
    return "\n".join(L)


def make_embed(path, assets, alt_map=None):
    """生成图片嵌入片段；给了 alt_map 就用当前模型写好的 alt，否则退回文件名。"""
    alt = os.path.splitext(os.path.basename(path))[0]
    if alt_map:
        alt = alt_map.get(rel_path(path, assets), alt)
    return "![%s](%s)" % (alt, rel_path(path, assets))


def extract_docx(src, assets, stem):
    import zipfile
    out = []
    with zipfile.ZipFile(src) as z:
        for i, m in enumerate(sorted(n for n in z.namelist() if n.startswith("word/media/")), 1):
            ext = m.rsplit(".", 1)[-1].lower()
            out.append(save_image(z.read(m), assets, stem, i, ext))
    return out


def extract_pptx(src, assets, stem):
    import zipfile
    out = []
    with zipfile.ZipFile(src) as z:
        for i, m in enumerate(sorted(n for n in z.namelist() if n.startswith("ppt/media/")), 1):
            ext = m.rsplit(".", 1)[-1].lower()
            out.append(save_image(z.read(m), assets, stem, i, ext))
    return out


def extract_pdf(src, assets, stem):
    import fitz
    out, idx = [], 0
    doc = fitz.open(src)
    for pi in range(doc.page_count):
        for img in doc.get_page_images(pi):
            idx += 1
            base = doc.extract_image(img[0])
            out.append(save_image(base["image"], assets, "%s_p%02d" % (stem, pi + 1), idx, base["ext"]))
    return out


def extract_html(src, assets, stem):
    from bs4 import BeautifulSoup
    import requests as req
    # 修复：使用 with 语句确保文件句柄正确关闭
    with open(src, encoding="utf-8", errors="ignore") as f:
        html = f.read()
    soup = BeautifulSoup(html, "html.parser")
    out, idx, base_dir = [], 0, os.path.dirname(os.path.abspath(src))
    for img in soup.find_all("img"):
        sa = img.get("src") or ""
        if not sa or sa.startswith("data:"):
            continue
        idx += 1
        if sa.startswith("http://") or sa.startswith("https://"):
            try:
                r = req.get(sa, timeout=20)
                data, ext = r.content, (r.headers.get("content-type", "").split("/")[-1] or "png").replace("jpeg", "jpg")
            except Exception:
                continue
        else:
            local = os.path.join(base_dir, sa)
            if not os.path.exists(local):
                continue
            with open(local, "rb") as f:
                data = f.read()
            ext = local.rsplit(".", 1)[-1].lower()
        out.append(save_image(data, assets, stem, idx, ext))
    return out


def extract_image(src, assets, stem):
    files = sorted(glob.glob(os.path.join(src, "*"))) if os.path.isdir(src) else [src]
    out, idx = [], 0
    for f in files:
        if os.path.splitext(f)[1].lower().lstrip(".") in IMG_EXT:
            idx += 1
            with open(f, "rb") as fh:
                out.append(save_image(fh.read(), assets, stem, idx, os.path.splitext(f)[1].lower().lstrip(".")))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fmt", choices=["docx", "pptx", "pdf", "html", "image"])
    ap.add_argument("src")
    ap.add_argument("--out", default=None)
    ap.add_argument("--recognize", action="store_true",
                    help="列出候选清单 + 打印作业单，由当前模型看图写 alt（不调网关）")
    ap.add_argument("--alt-file", default=None, metavar="清单文件",
                    help="按「相对路径<TAB>alt」清单生成嵌入片段（配合 --recognize 的作业单）")
    ap.add_argument("--split", action="store_true")
    ap.add_argument("--emit", default=None)
    a = ap.parse_args()
    stem = os.path.splitext(os.path.basename(a.src))[0]
    assets = assets_dir(a.src, a.out)
    exts = {"docx": extract_docx, "pptx": extract_pptx, "pdf": extract_pdf,
            "html": extract_html, "image": extract_image}
    raw = exts[a.fmt](a.src, assets, stem)
    targets = []
    for p in raw:
        targets.extend(split_by_whitespace(p) if a.split else [p])

    # ---- 只列候选 + 打印作业单（rc=2，待当前模型接手）----
    if a.recognize and not a.alt_file:
        cp = write_alt_candidates(assets, targets)
        sys.stderr.write(alt_handoff_text(assets, targets) + "\n")
        sys.stderr.write("[alt] 候选清单已写: %s\n" % cp)
        return 2

    alt_map = load_alt_map(a.alt_file) if a.alt_file else None
    embeds = [make_embed(t, assets, alt_map) for t in targets]
    text = "\n\n".join(embeds) + "\n"
    if a.emit:
        with open(a.emit, "w", encoding="utf-8") as f:
            f.write(text)
        print("[写入] %s (%d 张图)" % (a.emit, len(embeds)))
    else:
        sys.stdout.write(text)
    sys.stderr.write("[完成] 资源目录: %s，图片 %d 张\n" % (assets, len(embeds)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
