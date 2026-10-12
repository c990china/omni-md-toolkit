# -*- coding: utf-8 -*-
"""omni-md-toolkit 单文件转换入口：把任意文档转成干净、合规、可人工审阅的 Markdown。

特性：
  - 文本型：
      * docx: 本地精准合并单元格与图片落盘 (docx_merge_table)
      * doc: 自动经 Word COM 转为 docx 后走合并单元格修复 (doc_to_docx + docx_merge_table)
      * pptx: 形状排序与图片落盘 (docx_pptx_to_md)
      * pdf(文本型): 本地原生分栏聚类、编号定级、CID清洗直接抽取 (extract_text_pdf) 或 markitdown 引擎
      * xlsx/csv/html/msg/rtf/odt: markitdown 引擎（带 pandas/bs4 自动降级）
      * epub/azw3/mobi/kf8: 本地原生引擎 kindle_to_md（**不走 markitdown**，见TEXT_EXTS 注释）
  - 扫描型：扫描版 PDF / 纯图片
      * **不做 OCR，也不走任何网关/密钥**：本工具只负责渲染页面并逐页落盘，把图片
        交给**正在使用本技能的模型**自己看图写逐页稿（`_pages/<基名>/page_NN_ocr.md`）；
        全部写完后重跑本命令即自动合并。已存在的逐页稿自动跳过 ⇒ **支持断点续跑**。
        退出码 3 = 等待当前模型接手（不是失败）。
      * 若当前模型没有视觉能力：脚本会打印作业单提醒，请在客户端换用具备视觉能力的模型后重跑。
  - 后处理：版式自整理(md_reorganizer 最新修复版) -> 广告清理(ad_clean) -> 34条质量门禁(md_lint)
  - 产物：**只产出一个文件**，即与源文件同名的 `<基名>.md`（已自整理、已体检）；
    不再产出 `<基名>_整理.md` 之类的副本。

用法：
  python convert.py "a.docx"
  python convert.py "old.doc"                 # 自动转 docx 并修复表格
  python convert.py "book.pdf"                # 文本型 PDF 原生高保真抽取
  python convert.py "扫描件.pdf"                  # 渲染页面并打印作业单（当前模型看图）
  python convert.py "扫描件.pdf"                  # 逐页稿写好后重跑，即自动合并 + 自整理
  python convert.py "a.pdf" -o "a.md" --assets-dir DIR
  python convert.py "a.pdf" --no-clean        # 【仅调试】跳过自整理，产物为未整理稿，勿交付
  python convert.py "a.pdf" --no-gate         # 跳过质量门禁
  python convert.py "a.pdf" --source "a.docx" # 附带原件无损字符比对
"""
import os
import re
import sys
import argparse
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


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


TEXT_EXTS = {
    ".docx": "docx", ".doc": "doc", ".pptx": "pptx", ".html": "html", ".htm": "html",
    ".xlsx": "xlsx", ".xls": "xls", ".csv": "csv", ".msg": "msg",
    ".txt": "txt", ".md": "md", ".epub": "epub", ".rtf": "rtf", ".odt": "odt",
    # Kindle 电子书：走本地原生抽取引擎（kindle_to_md），**不能走 markitdown**。
    # 原因：markitdown 的 epub 分支依赖可选依赖 ebooklib，缺失时会静默退化为 zip 文件清单，
    # 正文一字不抽（2026-10-05 实测 16 章正文全丢，仅 1797 字符产物）。
    ".azw3": "kindle", ".mobi": "kindle", ".kf8": "kindle",
}
SCAN_EXTS = {".pdf": "pdf", ".png": "img", ".jpg": "img", ".jpeg": "img", ".webp": "img"}


def classify_pdf(path):
    """返回 'text' 或 'scan'"""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            rows = [(len(p.chars), len(p.images)) for p in pdf.pages]
        total_chars = sum(c for c, _ in rows)
        total_images = sum(im for _, im in rows)
        image_pages = sum(1 for c, im in rows if c == 0 or (im > 0 and c < 20))
        n = len(rows)
        if total_chars == 0 or image_pages >= max(1, n // 2):
            return "scan"
        if total_images > 0 and total_chars < total_images:
            return "scan"
        return "text"
    except Exception:
        try:
            import fitz
            doc = fitz.open(path)
            chars = sum(len(p.get_text().strip()) for p in doc)
            doc.close()
            return "text" if chars > 50 else "scan"
        except Exception:
            return "text"


def extract_base64_images(md, assets_dir, stem):
    import base64
    import re
    os.makedirs(assets_dir, exist_ok=True)
    rel_dir = os.path.basename(assets_dir)
    pat = re.compile(r"!\[([^\]]*)\]\((data:([^;]+);base64,([^)]+))\)")
    cnt = [0]

    def repl(m):
        alt, _, mime, b64 = m.group(1), m.group(2), m.group(3), m.group(4)
        try:
            data = base64.b64decode(b64)
        except Exception:
            return m.group(0)
        ext = (mime.split("/")[-1] or "png").replace("jpeg", "jpg")
        cnt[0] += 1
        name = "%s_img_%02d.%s" % (stem, cnt[0], ext)
        with open(os.path.join(assets_dir, name), "wb") as f:
            f.write(data)
        return "![%s](%s/%s)" % (alt or name, rel_dir, name)

    return pat.sub(repl, md)


def convert_text(path, out_path, assets_dir, pdf_engine="extract_text", keep_decor=False):
    ext = os.path.splitext(path)[1].lower()
    stem = os.path.splitext(os.path.basename(path))[0]
    base_dir = os.path.dirname(os.path.abspath(out_path))

    if ext == ".doc":
        doc_mod = _load("doc_to_docx")
        docx_path, ok = doc_mod.convert_doc(path, base_dir)
        if not ok or not os.path.exists(docx_path):
            raise RuntimeError(f".doc 转换失败: {path}")
        m = _load("docx_merge_table")
        return m.convert(docx_path, base_dir)

    if ext == ".docx":
        m = _load("docx_merge_table")
        return m.convert(path, base_dir)

    if ext == ".pptx":
        m = _load("docx_pptx_to_md")
        return m.convert_pptx(path, assets_dir, stem)

    if ext in (".azw3", ".mobi", ".kf8", ".epub"):
        # Kindle/电子书：本地原生解包 + DOM 抽取（markitdown 的 epub 分支不可靠，见 TEXT_EXTS 注释）
        m = _load("kindle_to_md")
        return m.extract(path, assets_dir, stem, keep_decor=keep_decor)

    if ext == ".pdf":
        if pdf_engine == "extract_text":
            try:
                m = _load("extract_text_pdf")
                return m.extract_text_to_markdown(path)
            except Exception as e:
                sys.stderr.write(f"[警告] extract_text_pdf 提取失败 ({e})，降级到 markitdown...\n")

    if ext in (".txt", ".md"):
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()

    # 其它格式（或 markitdown 模式 PDF）走 markitdown
    try:
        from markitdown import MarkItDown
        try:
            _res = MarkItDown().convert(path, keep_data_uris=True)
            md = getattr(_res, 'markdown', None) or _res.text_content
        except TypeError:
            _res = MarkItDown().convert(path)
            md = getattr(_res, 'markdown', None) or _res.text_content
        return extract_base64_images(md, assets_dir, stem)
    except ImportError:
        # markitdown 未安装时降级处理
        if ext in (".xlsx", ".xls"):
            import pandas as pd
            dfs = pd.read_excel(path, sheet_name=None)
            chunks = []
            for sheet_name, df in dfs.items():
                chunks.append(f"## {sheet_name}\n\n" + df.to_markdown(index=False))
            return "\n\n".join(chunks)
        elif ext == ".csv":
            import pandas as pd
            df = pd.read_csv(path)
            return df.to_markdown(index=False)
        else:
            raise RuntimeError(f"缺少依赖 markitdown，无法转换 {ext} 格式文件。请安装：pip install markitdown[all]")


class _NeedAI(Exception):
    """扫描件路线①：需要「当前模型」逐页看图写稿后再重跑。携带作业单文本。"""

    def __init__(self, src, plan):
        self.src = src
        self.plan = plan
        try:
            m = _load("ocr_llm")
            self.text = m.handoff_text(src, plan)
        except Exception:
            self.text = ("[视觉识别] 待识别 %d 页，页面目录：%s"
                         % (len(plan.get("todo", [])), plan.get("pages_dir")))
        super().__init__("需要当前模型逐页识别")


def convert_scan(path, out_path, assets_dir, pages_dir=None, dpi=200):
    """扫描型 / 图片型 → Markdown。

    **不做 OCR，也不调任何网关**：只渲染页面 + 逐页落盘，把作业单抛回调用方，
    由**正在使用本技能的模型**自己看图写逐页稿；齐了就直接合并。
    """
    m = _load("ocr_llm")
    plan = m.plan_pages(path, pages_dir, dpi)
    if plan["todo"]:
        raise _NeedAI(path, plan)
    md, missing = m.merge_drafts(plan["pages_dir"])
    if missing or not md.strip():
        raise _NeedAI(path, plan)
    return md


def postprocess(md, out_path, do_clean, ad_image):
    """版式自整理 → 广告清理。**任何交付产物都必须过这里**（除非显式 --no-clean）。"""
    if not do_clean:
        return md
    mr = _load("md_reorganizer")
    # ★ 2026-10-05：改用 `full_reorganize()`（CLI 步骤 0/7~7/7 的库函数版），
    #   而不是只调 `reorganize()` + `normalize_blank_lines()`。
    #   原因：完整流水线有 20+ 步（页眉清理、断裂标题合并、加粗子标题升级、
    #   错误标题降级、编号归一…），此前**只在 CLI 里写了、单文件路径漏了**，
    #   导致同一份 azw3：单文件入口 719 行 / 20 标题 / W0，
    #   批量流水线 2059 行 / 25 标题 / W5（目录项被误升成 5 个 `##`）。
    #   这正是 SKILL 踩坑 #14 的「同一逻辑写在两条路径必然漂移」，现已消除。
    md = mr.full_reorganize(md)
    ad = _load("ad_clean")
    md, _ = ad.clean_text(md)
    if ad_image:
        md, _ = ad.clean_image(md, os.path.dirname(os.path.abspath(out_path)))
    return md


def gate(md_path, src_path):
    ml = _load("md_lint")
    argv = [md_path]
    if src_path and os.path.exists(src_path):
        argv += ["--source", src_path]
    saved = sys.argv
    sys.argv = ["md_lint"] + argv
    try:
        rc = ml.main()
    except SystemExit as e:
        rc = e.code if isinstance(e.code, int) else 1
    finally:
        sys.argv = saved
    return rc


def run(src, out=None, assets_dir=None, do_clean=True, do_gate=True, ad_image=False, src_for_loss=None, pdf_engine="extract_text", pages_dir=None, dpi=200, keep_decor=False):
    ext = os.path.splitext(src)[1].lower()
    stem = os.path.splitext(os.path.basename(src))[0]
    out_path = out or (os.path.splitext(src)[0] + ".md")
    base_dir = os.path.dirname(os.path.abspath(out_path))
    # 资源目录名去掉尾随空格/点：Windows 不允许，否则目录能建但里面任何文件都写不进去
    # （Errno 2，只表现为「图片静默丢失」，极难排查；实测书名「目的与利润(精校本) .azw3」触发）
    assets_dir = assets_dir or os.path.join(
        base_dir, "assets_" + re.sub(r"[ .]+$", "", stem))
    os.makedirs(base_dir, exist_ok=True)
    # 图片资源目录必须在这里就建好：pptx 转换器（docx_pptx_to_md.convert_pptx）只负责写文件，
    # 不建目录；缺目录会让每张图片静默保存失败（仅 stderr 一条 warn），最终产出无图的 MD。
    os.makedirs(assets_dir, exist_ok=True)

    if ext in TEXT_EXTS and ext != ".pdf":
        kind = TEXT_EXTS[ext]
    elif ext == ".pdf":
        kind = classify_pdf(src)
        kind = "scan" if kind == "scan" else "pdf"
    elif ext in SCAN_EXTS:
        kind = "scan" if SCAN_EXTS[ext] == "pdf" else "img"
    else:
        sys.stderr.write("[错误] 不支持的扩展名：%s\n" % ext)
        return 2

    if kind in ("scan", "img"):
        route = "扫描型(由当前模型看图识别)"
    else:
        route = "文本型(%s)" % kind
    sys.stderr.write("[路由] %s -> %s\n" % (src, route))

    if kind in ("scan", "img"):
        try:
            md = convert_scan(src, out_path, assets_dir, pages_dir=pages_dir, dpi=dpi)
        except _NeedAI as e:
            sys.stderr.write(e.text + "\n\n")
            sys.stderr.write("[待接手] 本次未产出最终 MD —— 扫描件需要视觉能力。\n")
            sys.stderr.write("[待接手] 请按上面的作业单逐页识别，然后**重新运行本命令**，"
                             "即会自动合并并继续流水线。\n")
            return 3
    else:
        md = convert_text(src, out_path, assets_dir, pdf_engine=pdf_engine, keep_decor=keep_decor)

    # 自整理（默认开启）：产物**只写一次**，直接写与源文件同名的 <基名>.md。
    # 整理失败时绝不回落交付未整理稿——否则「产物必须已自整理」就形同虚设。
    clean_path = out_path
    if do_clean:
        try:
            md2 = postprocess(md, out_path, do_clean, ad_image)
        except Exception as e:
            sys.stderr.write("[错误] 版式自整理失败：%s\n" % e)
            sys.stderr.write("[错误] 为遵守「产物必须已自整理」，本次不写盘；请修复后重跑。\n")
            return 1
    else:
        md2 = md
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md2)
    sys.stderr.write("[%s] 已写: %s (%d 字符)\n"
                     % ("整理" if do_clean else "转换(未整理)", out_path, len(md2)))

    # 交付前自检：产物为空（无任何非空白字符）时绝不报「完成」——空文件没有使用价值。
    # 注意：--no-gate 会跳过体检，本检查不受该开关影响（0 字节产物一律阻断）。
    _delivered = md2
    if not _delivered.strip():
        sys.stderr.write("[错误] 转换产物无有效内容（%d 字符，均为空白）：%s\n"
                         % (len(_delivered), clean_path))
        sys.stderr.write("[错误] 可能原因：源文件本身无内容、该格式无可提取文字（扫描件需视觉识别）、"
                         "或解析器不支持该变体。\n")
        sys.stderr.write("[错误] 请检查源文件后重跑；本次不视为成功。\n")
        return 1

    gate_rc = 0
    if do_gate:
        sys.stderr.write("[门禁] 只读体检：\n")
        # ★ 电子书（azw3/mobi/kf8/epub）也纳入无损比对：md_lint.source_text 已支持
        #   这4 种格式（解包取纯正文作基准）。此前它们走 None⇒ **只过格式门禁、
        #   丢字完全不报警**（与 FAQ 10/11 同类的「静默不检查」，2026-10-05 审计发现）。
        loss_src = src_for_loss or (src if ext in (
            ".docx", ".doc", ".pdf", ".txt",
            ".azw3", ".mobi", ".kf8", ".epub") else None)
        gate_rc = gate(clean_path, loss_src)

    if gate_rc == 1:
        sys.stderr.write("[门禁拦截] 存在 E 级问题，产物未通过质检：%s\n" % clean_path)
        sys.stderr.write("[门禁拦截] 请先按上方清单修复后重跑；不要直接交付该文件。\n")
        return 1
    if gate_rc == 3:
        sys.stderr.write("[门禁不可信] md_lint 运行错误（rc=3），本次未完成质检：%s\n" % clean_path)
        sys.stderr.write("[门禁不可信] 请单独运行 md_lint.py 复核后再交付。\n")
        return 1

    sys.stderr.write("[完成] 交付：%s%s\n"
                     % (clean_path, "（仅有 W 级告警，未拦截）" if gate_rc == 2 else ""))
    return 0


def main():
    ap = argparse.ArgumentParser(description="omni-md-toolkit 单文件转换入口")
    ap.add_argument("src", help="源文件（docx/doc/pptx/pdf/xlsx/csv/html/msg/图片…）")
    ap.add_argument("-o", "--output", default=None, help="输出 .md（默认同名同目录）")
    ap.add_argument("--assets-dir", default=None, help="图片资源目录")
    ap.add_argument("--pdf-engine", choices=["extract_text", "markitdown"], default="extract_text", help="文本型 PDF 抽取引擎（默认 extract_text，排版更优）")
    ap.add_argument("--no-clean", action="store_true",
                    help="【仅调试】跳过版式自整理；产物为未整理稿，勿直接交付")
    ap.add_argument("--no-gate", action="store_true", help="跳过无损门禁")
    ap.add_argument("--ad-image", action="store_true", help="识别图片广告（由当前模型看图判断）")
    ap.add_argument("--pages-dir", default=None,
                    help="扫描件逐页工作目录（默认 <源文件目录>/_pages/<基名>/）")
    ap.add_argument("--dpi", type=int, default=200, help="扫描件渲染 DPI（默认 200）")
    ap.add_argument("--keep-decor", action="store_true",
                    help="电子书：保留章首装饰小图（默认剔除）")
    ap.add_argument("--source", default=None, help="原文件（用于字符级多重集比对）")
    a = ap.parse_args()

    return run(a.src, a.output, a.assets_dir, not a.no_clean, not a.no_gate, a.ad_image,
               a.source, pdf_engine=a.pdf_engine,
               pages_dir=a.pages_dir, dpi=a.dpi, keep_decor=a.keep_decor)


if __name__ == "__main__":
    sys.exit(main())
