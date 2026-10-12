# -*- coding: utf-8 -*-
"""DOCX / PPTX -> Markdown：按原文顺序输出文本与图片，图片原位内联嵌入。

特点：
  - 保留实际换行：块（段落/文本框/图片）之间空一行；段落内文本原样。
  - 图片内联：从文档抽取到 assets_<stem>/，按出现顺序命名，紧跟所在段落/形状插入。
  - DOCX 标题据样式(Heading/标题)转为 #；表格输出为基础 Markdown 表。
依赖：python-docx, python-pptx
用法：
  python docx_pptx_to_md.py docx "a.docx"        # 默认输出与源文件同名的 a.md（已自整理）
  python docx_pptx_to_md.py pptx "a.pptx" -o "a.md"
"""
import os, sys, io, re, argparse
# 用 reconfigure 而不是替换 sys.stdout：稍后加载 md_reorganizer 时它会再包一层
# io.TextIOWrapper，二次包装会让前一个包装器被回收时关闭底层缓冲，导致后续输出报错。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

IMG_EXT = ("png", "jpg", "jpeg", "gif", "bmp", "webp")


def save_image(data, assets, stem, idx, ext):
    name = "%s_%02d.%s" % (stem, idx, ext if ext in IMG_EXT else "png")
    path = os.path.join(assets, name)
    with open(path, "wb") as f:
        f.write(data)
    return path


def rel(path, base):
    return os.path.relpath(path, base or ".").replace(os.sep, "/")


def _para_text(p_element, qn):
    """段落文本：把段落内硬换行(br/cr)与制表符合并为空格，形成自然段落流。"""
    parts = []
    for node in p_element.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            parts.append(" ")
        elif node.tag == qn("w:tab"):
            parts.append(" ")
    return re.sub(r" {2,}", " ", "".join(parts)).strip()


def convert_docx(src, assets, stem):
    import docx
    from docx.oxml.ns import qn
    from collections import Counter
    d = docx.Document(src)
    body = d.element.body
    SENT_END = "。！？.!?」』）】"
    PAGE_RE = re.compile(r"^.{0,6}第\s*\d+\s*页.{0,6}$")
    # 第一遍：收集段落（文本 / 样式 / 图片）
    items = []
    idx = 0
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            para = docx.text.paragraph.Paragraph(child, d)
            text = _para_text(child, qn)
            style = (para.style.name if para.style else "") or ""
            is_heading = style.startswith("Heading") or style.startswith("标题")
            lvl = 1
            if is_heading:
                m = "".join(filter(str.isdigit, style))
                lvl = int(m) if m else 1
            imgs = []
            for blip in child.iter(qn("a:blip")):
                rId = blip.get(qn("r:embed"))
                if not rId:
                    continue
                try:
                    ip = d.part.related_parts[rId]
                    idx += 1
                    ext = (ip.content_type.split("/")[-1] or "png").replace("jpeg", "jpg")
                    pp = save_image(ip.blob, assets, stem, idx, ext)
                    imgs.append(rel(pp, os.path.dirname(assets)))
                except Exception as e:
                    sys.stderr.write("[warn] docx 图片失败: %s\n" % e)
            items.append({"text": text, "heading": is_heading, "lvl": lvl,
                          "imgs": imgs, "empty": text.strip() == ""})
        elif child.tag == qn("w:tbl"):
            tbl = docx.table.Table(child, d)
            rows = []
            for row in tbl.rows:
                cells = [c.text.replace("\n", " ").replace("|", "/") for c in row.cells]
                rows.append("| " + " | ".join(cells) + " |")
            items.append({"table": rows})
    # 重复文本（页眉/页脚）计数，用于去噪
    cnt = Counter(it["text"].strip() for it in items
                 if "text" in it and it["text"].strip())
    out = []
    buf = ""
    def flush():
        if buf.strip():
            out.append(buf.strip())
        return ""
    for it in items:
        if "table" in it:
            buf = flush()
            out.extend(it["table"])
            out.append("")
            continue
        text = it["text"].strip()
        imgs = it["imgs"]
        # 跳过页眉/页脚（重复出现）与页码标记
        if text and (cnt[text] >= 10 or PAGE_RE.match(text)):
            for pp in imgs:
                out.append("![%s](%s)" % (os.path.splitext(os.path.basename(pp))[0], pp))
            continue
        if it["empty"]:
            buf = flush()
            for pp in imgs:
                out.append("![%s](%s)" % (os.path.splitext(os.path.basename(pp))[0], pp))
            continue
        if it["heading"]:
            buf = flush()
            out.append("%s %s" % ("#" * min(max(it["lvl"], 1), 6), text))
            out.append("")
            for pp in imgs:
                out.append("![%s](%s)" % (os.path.splitext(os.path.basename(pp))[0], pp))
            continue
        # 普通段落：句末标点处断段，否则与上行合并为自然段落流
        if buf and buf[-1] in SENT_END:
            buf = flush()
        buf += text
        for pp in imgs:
            buf += "\n\n![%s](%s)" % (os.path.splitext(os.path.basename(pp))[0], pp)
    buf = flush()
    return "\n".join(out).strip() + "\n"


def convert_pptx(src, assets, stem):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(src)
    idx = 0
    out = []
    for si, slide in enumerate(prs.slides, 1):
        out.append("## 第 %d 页" % si)
        shapes = sorted(slide.shapes, key=lambda s: (s.top or 0, s.left or 0))
        for sh in shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    img = sh.image
                    idx += 1
                    p = save_image(img.blob, assets, stem, idx, img.ext or "png")
                    out.append("![%s](%s)" % (os.path.splitext(os.path.basename(p))[0], rel(p, os.path.dirname(assets))))
                    out.append("")
                except Exception as e:
                    sys.stderr.write("[warn] pptx 图片失败: %s\n" % e)
            elif sh.has_text_frame and sh.text_frame.text.strip():
                # 本模块契约：块（段落/文本框/图片）之间空一行。缺这一行会让相邻文本框在
                # Markdown 里被并成同一个段落（单换行不分段），进而被整理器误判标题。
                out.append(sh.text_frame.text)
                out.append("")
        # 演讲者备注同样是幻灯片内容的一部分，不抽取就等于静默丢稿。
        # 用引用块与正文区分，避免与幻灯片正文混淆。
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            out.append("> **演讲者备注**：%s" % " ".join(notes.split()))
            out.append("")
        out.append("")
    return "\n".join(out).strip() + "\n"


def _load(name):
    """按「长文件名前缀模糊匹配」动态加载同目录脚本（禁止硬编码短名）。"""
    import importlib.util as _ilu
    here = os.path.dirname(os.path.abspath(__file__))
    for f in os.listdir(here):
        if f.endswith(name + ".py") or f.endswith("_" + name + ".py"):
            target = os.path.join(here, f)
            break
    else:
        target = os.path.join(here, name + ".py")
    spec = _ilu.spec_from_file_location(name, target)
    mod = _ilu.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fmt", choices=["docx", "pptx"])
    ap.add_argument("src")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 .md（默认与源文件同名的 <基名>.md）")
    ap.add_argument("--no-clean", action="store_true",
                    help="【仅调试】跳过版式自整理；产物为未整理稿，勿直接交付")
    a = ap.parse_args()
    if a.no_clean:
        sys.stderr.write("[警告] --no-clean 仅用于调试：本次产物为**未整理**稿，请勿直接交付。\n")
    stem = os.path.splitext(os.path.basename(a.src))[0]
    out = a.output or (os.path.splitext(a.src)[0] + ".md")
    assets = os.path.join(os.path.dirname(os.path.abspath(out)), "assets_" + stem)
    os.makedirs(assets, exist_ok=True)
    if a.fmt == "docx":
        md = convert_docx(a.src, assets, stem)
    else:
        md = convert_pptx(a.src, assets, stem)
    # 交付产物必须已自整理（convert.py 调用的是上面的转换函数，不会重复整理）。
    if not a.no_clean:
        try:
            md = _load("md_reorganizer").reorganize(md)
        except Exception as e:
            sys.stderr.write("[错误] 版式自整理失败：%s\n" % e)
            sys.stderr.write("[错误] 为遵守「产物必须已自整理」，本次不写盘；请修复后重跑。\n")
            return 1
    with open(out, "w", encoding="utf-8") as f:
        f.write(md)
    n_img = md.count("![")
    sys.stderr.write("[完成] %s -> %s （图片内联 %d 张，资源 %s）\n" % (a.src, out, n_img, assets))
    return 0


if __name__ == "__main__":
    sys.exit(main())
