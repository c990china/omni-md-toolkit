# -*- coding: utf-8 -*-
"""广告清理（文本广告 + 图片广告），作为 convert.py 后处理的一环。

设计要点：
  - clean_text(md): 按关键词删除「文字型」广告行（图片行不动），并重连段落、压缩空行。无需模型。
  - clean_image(md, base_dir, ad_list=...): 删除「图片型」广告/推广页并移除其引用。
      * **不做 OCR，也不调任何网关**：由**当前模型**看图判断，把广告图相对路径清单
        传给 ad_list / --ad-apply 落地删除。
      * 缺清单时脚本只列出候选清单并打印作业单，**不再静默跳过**。

参考来源：all-format-md-toolkit 的 clean_ad_text.py / filter_ads.py（已去除 toolkit_common 依赖，改为函数式 API）。
"""
import os
import re
import sys
import io
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
# 本脚本的真实路径（带中文前缀的长文件名）。作业单里必须打印它，
# 不能写短名 `ad_clean.py`——磁盘上并不存在该短名文件，照做会 FileNotFoundError。
_SELF_PATH = os.path.abspath(__file__)

DEFAULT_KW = [
    "精品课程", "课程售价", "立即购买", "扫码购买", "卖课",
    "关注公众号", "二维码", "扫码关注", "购买课程",
    "限时优惠", "加微信", "天成恒通",
]
IMG_RE = re.compile(r'^\s*!\[[^\]]*\]\(([^)]+)\)\s*$')


def _find_script(name):
    for f in os.listdir(HERE):
        if f.endswith(name + ".py") or f.endswith("_" + name + ".py"):
            return os.path.join(HERE, f)
    return os.path.join(HERE, name + ".py")


# ---------------------------------------------------------------- 文字型广告

def clean_text(md, keywords=None):
    """按关键词删除文字型广告行。返回 (清理后文本, 删除行数)。

    只处理「文字型广告」：命中关键词的非图片行会被删除，并轻量重连段落、压缩多余空行。
    图片型广告请用 clean_image（需视觉模型）或人工打开 assets 目录剔除。
    """
    kws = keywords or list(DEFAULT_KW)
    lines = md.split("\n")
    remove_idx = set()
    hits = []
    for i, ln in enumerate(lines):
        if IMG_RE.match(ln):
            continue  # 图片行交给 clean_image，这里不动
        low = ln.lower()
        for kw in kws:
            if kw.lower() in low:
                remove_idx.add(i)
                hits.append((i + 1, kw, ln.strip()))
                break
    if not hits:
        return md, 0
    new_lines = [ln for i, ln in enumerate(lines) if i not in remove_idx]
    out, blank = [], 0
    for ln in new_lines:
        if ln.strip() == "":
            blank += 1
            if blank <= 1:
                out.append(ln)
        else:
            blank = 0
            out.append(ln)
    sys.stderr.write("[广告] 文字型广告命中 %d 行，已删除\n" % len(hits))
    return "\n".join(out).strip() + "\n", len(hits)


# ---------------------------------------------------------------- 图片型广告

CANDIDATES_NAME = "_ad_candidates.txt"
RESULT_NAME = "_ad_result.txt"


def list_images(md, base_dir):
    """列出 md 里引用的全部图片。返回 [(行号, 相对路径, 绝对路径)]。"""
    refs = []
    for i, ln in enumerate(md.split("\n")):
        m = IMG_RE.match(ln)
        if m:
            rel = m.group(1)
            refs.append((i, rel, os.path.normpath(os.path.join(base_dir, rel))))
    return refs


def _norm_key(s):
    return os.path.normpath(str(s)).replace("\\", "/")


def _remove_refs(md, refs, picked, dry_run=False):
    """按 picked=[(行号, 相对路径, 绝对路径)] 删除图片及其引用行。返回 (新文本, 删除张数)。"""
    remove_idx = set(i for i, _r, _p in picked)
    removed = [r for _i, r, _p in picked]
    kept = [r for i, r, _p in refs if i not in remove_idx]

    if dry_run:
        sys.stderr.write("[广告] 将删除 %d 张广告图，保留 %d 张（dry-run，不删）\n"
                         % (len(removed), len(kept)))
        return md, len(removed)

    for _i, _rel, apath in picked:
        try:
            if os.path.exists(apath):
                os.remove(apath)
        except Exception as e:
            sys.stderr.write("[warn] 删除失败 %s: %s\n" % (apath, e))

    lines = md.split("\n")
    new_lines = [ln for i, ln in enumerate(lines) if i not in remove_idx]
    out, blank = [], 0
    for ln in new_lines:
        if ln.strip() == "":
            blank += 1
            if blank <= 1:
                out.append(ln)
        else:
            blank = 0
            out.append(ln)
    sys.stderr.write("[广告] 删除广告图 %d 张，保留 %d 张\n" % (len(removed), len(kept)))
    return "\n".join(out).strip() + "\n", len(removed)


def apply_ad_list(md, base_dir, ad_paths, dry_run=False):
    """按给定的「广告图相对路径清单」删除图片及其引用。返回 (新文本, 删除张数)。

    这是路线①（当前模型看图判断）的落地函数：不调任何网关。
    """
    refs = list_images(md, base_dir)
    if not refs:
        return md, 0
    want = set(_norm_key(x) for x in ad_paths if str(x).strip())
    picked = [(i, rel, ap) for i, rel, ap in refs if _norm_key(rel) in want]
    miss = want - set(_norm_key(r) for _i, r, _a in refs)
    if miss:
        sys.stderr.write("[广告] 清单里有 %d 条在本文档找不到，已忽略：%s\n"
                         % (len(miss), ", ".join(sorted(miss)[:5])))
    if not picked:
        sys.stderr.write("[广告] 清单里没有任何图片命中本文档，未改动。\n")
        return md, 0
    return _remove_refs(md, refs, picked, dry_run)


def write_candidates(base_dir, refs, path=None):
    """把候选图片清单写到文件（每行一个相对路径），返回文件路径。"""
    p = path or os.path.join(base_dir, CANDIDATES_NAME)
    with io.open(p, "w", encoding="utf-8", newline="") as f:
        f.write("# 候选图片清单（ad_clean 生成，仅参考）。\n")
        f.write("# 请把**判定为广告**的图片相对路径另存为 %s 后运行：--ad-apply <文件>\n" % RESULT_NAME)
        for _i, rel, _ap in refs:
            f.write(rel + "\n")
    return p


def ad_handoff_text(refs, base_dir, md_path=None):
    """生成「请当前模型看图判断哪些图是广告」的作业单。"""
    result = os.path.join(base_dir, RESULT_NAME)
    L = []
    L.append("[图片广告] 需要判断下列 %d 张图片是否为广告/推广图。" % len(refs))
    if md_path:
        L.append("  目标 MD: %s" % os.path.abspath(md_path))
    L.append("")
    L.append("  ▶ 请由你（当前模型）逐张读取下面的图片，把**判定为广告**的相对路径写入：")
    L.append("       %s" % result)
    L.append("     （每行一个，如 `assets_x/xxx.png`；非广告的不要写）")
    L.append("     然后运行：")
    L.append('       python "%s" "%s" --ad-apply "%s"'
             % (_SELF_PATH, md_path or "<目标.md>", result))
    L.append("")
    L.append("  ▶ 本技能**不做 OCR，也不调任何网关**：判断一律由具备视觉能力的模型完成。")
    L.append("     若你（当前模型）没有视觉能力：请使用者提供一个视觉模型的")
    L.append("     【模型名 / 网关地址 / 密钥】，由你自行调用它完成判断（本工具不做 OCR、")
    L.append("     不含网关调用代码）；或切换到具备视觉能力的模型后重跑。")
    L.append("     或切换到具备视觉能力的模型后重跑。")
    L.append("")
    L.append("  候选图片：")
    for _i, rel, ap in refs[:50]:
        L.append("   - %s   (%s)" % (rel, ap))
    if len(refs) > 50:
        L.append("   ...（共 %d 张，此处仅列前 50）" % len(refs))
    return "\n".join(L)


def clean_image(md, base_dir, dry_run=False, mock=False, ad_list=None):
    """删除广告图片及其引用。返回 (清理后文本, 删除张数)。

    **不做 OCR，也不调任何网关**。优先级：ad_list > mock > 列候选清单 + 打印作业单。
      ad_list: 当前模型看图后给出的广告图相对路径集合（推荐；--ad-apply 落地）。
      mock:    全部判为广告（本地验证删除逻辑，不联网，会真删！务必配 dry_run 先预览）。
    """
    refs = list_images(md, base_dir)
    if not refs:
        return md, 0

    if ad_list is not None:
        return apply_ad_list(md, base_dir, ad_list, dry_run)

    if mock:
        return _remove_refs(md, refs, [(i, rel, ap) for i, rel, ap in refs], dry_run)

    # 无清单：只列候选 + 打印作业单（由当前模型看图判断），绝不静默跳过、也不调网关。
    cp = write_candidates(base_dir, refs)
    sys.stderr.write("[广告] 已列出 %d 张候选图片 -> %s\n" % (len(refs), cp))
    sys.stderr.write('[广告] 图片广告需**当前模型**看图判断；运行 "%s" --ad-image 可获取作业单。\n'
                     % _SELF_PATH)
    return md, 0


# ---------------------------------------------------------------- 命令行（可选单独使用）

def main():
    ap = argparse.ArgumentParser(description="广告清理（文字 + 图片）")
    ap.add_argument("path", help="单个 .md 文件，或包含 .md 的目录")
    ap.add_argument("--keywords", default="", help="自定义文字关键词，英文逗号分隔；省略用内置词表")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不删除")
    ap.add_argument("-r", "--recursive", action="store_true", help="目录模式递归子目录")
    ap.add_argument("--ad-image", action="store_true",
                    help="图片广告清理（由当前模型看图判断；清单落地用 --ad-apply）")
    ap.add_argument("--ad-apply", default=None, metavar="清单文件",
                    help="按清单删除广告图（每行一个图片相对路径）；配合 --ad-image 的作业单使用")
    ap.add_argument("--mock", action="store_true",
                    help="本地试跑：全部判为广告（验证删除逻辑，不联网，会真删！配 --dry-run 先预览）")
    a = ap.parse_args()

    import glob
    targets = []
    if os.path.isfile(a.path):
        targets = [os.path.abspath(a.path)]
    elif os.path.isdir(a.path):
        pat = os.path.join(a.path, "**", "*.md") if a.recursive else os.path.join(a.path, "*.md")
        targets = sorted(glob.glob(pat, recursive=a.recursive))
    else:
        sys.stderr.write("[错误] 路径不存在: %s\n" % a.path)
        sys.exit(2)
    if not targets:
        sys.stderr.write("[完成] 未找到需要处理的 .md 文件。\n")
        return

    # ---- 路线①之落地：按清单删除广告图 ----
    if a.ad_apply:
        ad_paths = [l.strip() for l in io.open(a.ad_apply, encoding="utf-8-sig")
                    if l.strip() and not l.strip().startswith("#")]
        n_all = 0
        for md_path in targets:
            text = io.open(md_path, encoding="utf-8-sig").read()
            base_dir = os.path.dirname(os.path.abspath(md_path))
            text, k = apply_ad_list(text, base_dir, ad_paths, dry_run=a.dry_run)
            n_all += k
            if not a.dry_run:
                io.open(md_path, "w", encoding="utf-8", newline="").write(text)
        sys.stderr.write("[广告] 清单 %d 条，实际删除 %d 张图。%s\n"
                         % (len(ad_paths), n_all, "（dry-run 未写入）" if a.dry_run else ""))
        return 0

    kws = [k.strip() for k in a.keywords.split(",") if k.strip()] if a.keywords else list(DEFAULT_KW)

    # ---- 输出候选清单 + 作业单（rc=2，待当前模型接手）----
    # 本技能不做 OCR、不调网关：图片广告一律由具备视觉能力的模型看图判断。
    if a.ad_image and not a.mock:
        for md_path in targets:
            base_dir = os.path.dirname(os.path.abspath(md_path))
            text = io.open(md_path, encoding="utf-8-sig").read()
            refs = list_images(text, base_dir)
            if not refs:
                sys.stderr.write("[广告] %s：无图片引用，跳过。\n" % os.path.basename(md_path))
                continue
            cp = write_candidates(base_dir, refs)
            sys.stderr.write(ad_handoff_text(refs, base_dir, md_path) + "\n")
            sys.stderr.write("[广告] 候选清单已写: %s\n" % cp)
        return 2
    total_txt, total_img = 0, 0
    for md_path in targets:
        sys.stderr.write("========== %s ==========\n" % md_path)
        with io.open(md_path, encoding="utf-8") as f:
            text = f.read()
        text, n_txt = clean_text(text, kws)
        total_txt += n_txt
        if a.ad_image:
            base_dir = os.path.dirname(os.path.abspath(md_path))
            text, n_img = clean_image(text, base_dir, dry_run=a.dry_run, mock=a.mock)
            total_img += n_img
        if not a.dry_run:
            with io.open(md_path, "w", encoding="utf-8") as f:
                f.write(text)
    sys.stderr.write("[汇总] 删除文字广告 %d 行，图片广告 %d 张\n" % (total_txt, total_img))


if __name__ == "__main__":
    sys.exit(main())
