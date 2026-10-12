# -*- coding: utf-8 -*-
"""扫描版 / 图片型 PDF 与图片 → Markdown：渲染页面 + 交给「当前模型」看图识别。

设计要点：
  - **不做 OCR，也不走任何网关 / 密钥**：本脚本只负责把页面渲染成图片并逐页落盘，
    然后把图片交给**正在使用本技能的模型**（当前模型）自己看图转 Markdown。
  - 逐页稿写到 `_pages/<源文件基名>/page_NN_ocr.md`；已存在的逐页稿自动跳过
    ⇒ 天然支持**断点续跑**。全部就绪后重跑本命令（或加 `--merge`）即自动合并。
  - 合并后的整篇 MD **会自动过一遍版式自整理**（md_reorganizer），并按「与源文件同名」
    的规则写成 `<源文件基名>.md`（**不再产出 `_整理.md` 之类的副本**）。
  - 若当前模型没有视觉能力，脚本会打印作业单提醒：请在客户端换用具备视觉能力的模型后重跑。
依赖：PyMuPDF(fitz)（渲染页面）。**不需要 requests**。
"""
import os
import sys
import argparse
import importlib.util

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))


def _find_script(name):
    for f in os.listdir(HERE):
        if f.endswith(name + ".py") or f.endswith("_" + name + ".py"):
            return os.path.join(HERE, f)
    return os.path.join(HERE, name + ".py")


def _load(name):
    target = _find_script(name)
    spec = importlib.util.spec_from_file_location(name, target)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ================= 渲染 + 交给当前模型（零网关、零密钥）=================

PAGES_DIRNAME = "_pages"
MANIFEST_NAME = "manifest.json"


def pages_dir_for(src_path, pages_dir=None):
    """逐页工作目录：默认 `<源文件同目录>/_pages/<源文件基名>/`。

    ★ 按源文件基名再分一层，避免**同一目录下的多个源文件共用页稿目录**而互相覆盖
      （例如 a.pdf 与 a.png 都用 page_01_ocr.md 就会串）。
    """
    if pages_dir:
        return pages_dir
    stem = os.path.splitext(os.path.basename(src_path))[0]
    return os.path.join(os.path.dirname(os.path.abspath(src_path)), PAGES_DIRNAME, stem)


def page_png_path(pages_dir, idx, ext=".png"):
    """第 idx 页的页面图片路径。

    ★ 单图输入时保留**源文件真实扩展名**（.jpg/.webp 不能硬叫 .png），
      否则文件名与内容不符，会误导看图的一方按错误格式解读。
    """
    return os.path.join(pages_dir, "page_%02d%s" % (idx, ext))


def page_draft_path(pages_dir, idx):
    """逐页稿路径：把第 idx 页的识别结果写到这里。"""
    return os.path.join(pages_dir, "page_%02d_ocr.md" % idx)


def render_pages(src_path, pages_dir=None, dpi=200):
    """把源文件渲染为逐页图片（PDF 逐页 PNG；单图 1 页，保留原扩展名），

    返回 [(idx, 页面图片路径)]。已存在的图片会**复用**（不重复渲染），便于断点续跑。
    """
    ext = os.path.splitext(src_path)[1].lower()
    pages_dir = pages_dir_for(src_path, pages_dir)
    os.makedirs(pages_dir, exist_ok=True)

    if ext != ".pdf":
        dst = page_png_path(pages_dir, 1, ext)   # 保留真实扩展名（.jpg/.webp…）
        if not os.path.exists(dst):
            import shutil
            shutil.copyfile(src_path, dst)
        return [(1, dst)]

    try:
        import fitz
    except Exception:
        raise RuntimeError("渲染 PDF 需要 PyMuPDF：pip install PyMuPDF")
    zoom = dpi / 72.0
    doc = fitz.open(src_path)
    pages = []
    for i, page in enumerate(doc):
        idx = i + 1
        png = page_png_path(pages_dir, idx)
        if not os.path.exists(png):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            pix.save(png)
        pages.append((idx, png))
    doc.close()
    return pages


def plan_pages(src_path, pages_dir=None, dpi=200):
    """渲染 + 统计进度，并写 manifest.json。

    返回 dict: {pages, todo(未识别页号), done(已有稿页号), pages_dir, manifest}
    """
    pages_dir = pages_dir_for(src_path, pages_dir)
    pages = render_pages(src_path, pages_dir, dpi)
    todo, done = [], []
    for idx, _png in pages:
        (done if os.path.exists(page_draft_path(pages_dir, idx)) else todo).append(idx)

    manifest = {
        "src": os.path.abspath(src_path),
        "pages_dir": os.path.abspath(pages_dir),
        "total": len(pages),
        "pages": [{"page": idx, "png": os.path.abspath(png),
                   "draft": os.path.abspath(page_draft_path(pages_dir, idx))}
                  for idx, png in pages],
    }
    try:
        import json
        with open(os.path.join(pages_dir, MANIFEST_NAME), "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return {"pages": pages, "todo": todo, "done": done,
            "pages_dir": pages_dir, "manifest": manifest}


def merge_drafts(pages_dir, out_path=None):
    """从磁盘逐页稿重建合并稿（**不依赖内存**）。返回 (md_text, missing_idx_list)。

    缺页**不静默**：missing 里逐页列出，由调用方显式处理。
    """
    import json
    idxs = []
    mpath = os.path.join(pages_dir, MANIFEST_NAME)
    if os.path.exists(mpath):
        try:
            m = json.load(open(mpath, encoding="utf-8"))
            idxs = [int(x["page"]) for x in m.get("pages", [])]
        except Exception:
            idxs = []
    if not idxs:
        import glob as _glob
        import re as _re
        for f in sorted(_glob.glob(os.path.join(pages_dir, "page_*_ocr.md"))):
            mm = _re.search(r"page_(\d+)_ocr\.md$", os.path.basename(f))
            if mm:
                idxs.append(int(mm.group(1)))
    idxs = sorted(set(idxs))

    blocks, missing = [], []
    for idx in idxs:
        dp = page_draft_path(pages_dir, idx)
        if os.path.exists(dp):
            t = open(dp, encoding="utf-8-sig").read().strip()
            blocks.append("<!-- 第 %d 页 -->" % idx)
            blocks.append(t)
            blocks.append("")
        else:
            missing.append(idx)
    md = "\n".join(blocks).strip()
    md = (md + "\n") if md else ""
    if out_path and md:
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            f.write(md)
    return md, missing


def organize(md):
    """过一遍版式自整理（md_reorganizer 最新修复版）。"""
    mr = _load("md_reorganizer")
    return mr.reorganize(md)


def default_out_path(src_path):
    """默认产物：与源文件同名的 `<源文件基名>.md`（同目录）。"""
    return os.path.splitext(os.path.abspath(src_path))[0] + ".md"


def handoff_text(src_path, plan):
    """生成给「当前模型」的作业单（纯文本，便于直接照做）。"""
    pd = plan["pages_dir"]
    L = []
    L.append("[视觉识别] 扫描页 → Markdown")
    L.append("  源文件: %s" % os.path.abspath(src_path))
    L.append("  页面图片目录: %s" % os.path.abspath(pd))
    L.append("  总页数 %d ；已识别 %d ；待识别 %d"
             % (len(plan["pages"]), len(plan["done"]), len(plan["todo"])))
    L.append("")
    L.append("  ▶ 请由你（当前模型）逐页读取下面的页面图片，把该页 Markdown 写入对应的 _ocr.md：")
    _img_of = {i: p for i, p in plan["pages"]}
    for idx in plan["todo"][:50]:
        L.append("       第 %d 页: 读 %s  ->  写 %s"
                 % (idx, _img_of.get(idx) or page_png_path(pd, idx),
                    page_draft_path(pd, idx)))
    if len(plan["todo"]) > 50:
        L.append("       ...（共 %d 页待识别，此处仅列前 50）" % len(plan["todo"]))
    L.append("     全部写完后重新运行本命令（或加 --merge）即可合并成整篇。")
    L.append("     合并时会自动过一遍版式自整理，产物为与源文件同名的 <基名>.md。")
    L.append("     ★ 已存在的逐页稿会自动跳过 ⇒ 中途中断可从断点继续，不会重做已完成页。")
    L.append("")
    L.append("  ▶ 本技能**不做 OCR，也不调用任何网关**：识别一律由具备视觉能力的模型完成。")
    L.append("     若你（当前模型）没有视觉能力：请使用者提供一个视觉模型的")
    L.append("     【模型名 / 网关地址 / 密钥】，由你自行调用它完成识别（本工具不做 OCR、")
    L.append("     不含网关调用代码）；或切换到具备视觉能力的模型后重跑。")
    L.append("     或切换到具备视觉能力的模型后重跑本命令。")
    return "\n".join(L)


def _do_merge(a, pdir):
    """把逐页稿合并成整篇 → 自整理 → 写成与源文件同名的 .md。返回进程码。"""
    out_path = a.output or default_out_path(a.src)
    no_clean = getattr(a, "no_clean", False)
    if no_clean:
        sys.stderr.write("[警告] --no-clean 仅用于调试：本次产物为**未整理**稿，请勿直接交付。\n")
    md, missing = merge_drafts(pdir)
    if missing:
        sys.stderr.write("[错误] 以下页缺逐页稿，未合并：%s\n" % ", ".join(map(str, missing)))
        sys.stderr.write("[错误] 请补齐后重跑；缺页不会被静默跳过。\n")
        return 1
    if not md.strip():
        sys.stderr.write("[错误] 逐页稿全为空，合并结果为空（%s）。\n" % pdir)
        return 1
    # 交付产物必须已自整理；整理失败时宁可报错不产出，也不交付未整理稿。
    if not no_clean:
        try:
            md = organize(md)
        except Exception as e:
            sys.stderr.write("[错误] 版式自整理失败：%s\n" % e)
            sys.stderr.write("[错误] 为遵守「产物必须已自整理」，本次不写盘；请修复后重跑。\n")
            return 1
    try:
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            f.write(md)
    except Exception as e:
        sys.stderr.write("[错误] 写入失败 %s: %s\n" % (out_path, e))
        return 1
    sys.stderr.write("[完成] 已写%s: %s\n"
                     % ("（已自整理）" if not no_clean else "（未整理）", out_path))
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="扫描/图片 → Markdown（由当前模型看图识别；不做 OCR、不走网关）")
    ap.add_argument("src", help="扫描型 PDF 或图片（png/jpg/jpeg/webp）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 .md（默认与源文件同名的 <基名>.md）")
    ap.add_argument("--pages-dir", default=None,
                    help="逐页工作目录（默认 <源文件目录>/_pages/<基名>/）")
    ap.add_argument("--dpi", type=int, default=200, help="PDF 渲染 DPI（默认 200）")
    ap.add_argument("--render-only", action="store_true",
                    help="只渲染页面并打印作业单（交给当前模型看图），不做合并")
    ap.add_argument("--merge", action="store_true",
                    help="把已落盘的逐页稿合并成整篇（自整理后写成 <基名>.md）")
    ap.add_argument("--no-clean", action="store_true",
                    help="【仅调试】合并时不套用版式自整理；产物为未整理稿，勿直接交付")
    a = ap.parse_args()

    pdir = pages_dir_for(a.src, a.pages_dir)

    # ---- 显式合并 ----
    if a.merge:
        return _do_merge(a, pdir)

    # ---- 只渲染 + 交作业 ----
    if a.render_only:
        plan = plan_pages(a.src, a.pages_dir, a.dpi)
        if not plan["todo"]:
            sys.stderr.write("[视觉识别] 全部 %d 页逐页稿已就绪。\n" % len(plan["pages"]))
            return 0
        sys.stderr.write(handoff_text(a.src, plan) + "\n")
        return 2

    # ---- 默认：渲染 + 交作业；逐页稿已齐则直接合并 ----
    plan = plan_pages(a.src, a.pages_dir, a.dpi)
    if not plan["todo"]:
        sys.stderr.write("[视觉识别] 全部 %d 页逐页稿已就绪，直接合并。\n" % len(plan["pages"]))
        return _do_merge(a, pdir)
    sys.stderr.write(handoff_text(a.src, plan) + "\n")
    return 2          # 2 = 待当前模型接手（不是失败）


if __name__ == "__main__":
    sys.exit(main())
