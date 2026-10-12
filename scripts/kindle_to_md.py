# -*- coding: utf-8 -*-
"""omni-md-toolkit 补充引擎：Kindle 电子书（.azw3 / .mobi / .kf8 / .epub）原生抽取引擎。

为什么需要本模块（2026-10-05 新增）：
  markitdown 对 `.epub` 的支持**依赖可选依赖 `ebooklib`**。实测环境markitdown 0.0.2
  未装 ebooklib 时，epub 会静默走 zip 回退分支，产物只有「## File: xxx」的文件清单、
  **正文一字不抽**（实测 16 章正文全丢，产物仅 1797 字符）。
  ⇒ 电子书必须走本模块的原生解析，不能依赖 markitdown 的 epub 分支。

能力：
  * `.azw3` / `.mobi` / `.kf8`：用 `mobi` 库（mobiunpack 的 Python 封装）解出 KF8 结构，
    优先取 `mobi8/OEBPS`，无 KF8 时回退 `mobi7` 的 MOBI HTML 流。
  * `.epub`：直接按 zip + content.opf + spine 顺序解析，零额外依赖。
  * 章节顺序按 **content.opf 的 <spine>** 排列（不靠文件名排序，避免 part1/part10 乱序）。
  * 标题层级还原：calibre 转换器把样式写进 class，本引擎按 class 语义还原H1/H2/H3；
    同时保留**通用兜底**（h1~h6 / <title> / 目录 navPoint），非 calibre 排版也能用。
  * 装饰性小图（章首 101x101 灰度图标等）默认**剔除**（`--keep-decor` 可保留）；
    封面与含二维码/推广信息的整页彩图保留并落盘。
  * 图片统一落盘到 `assets_<基名>/`，Markdown 内写**相对路径**，不写 base64。

库函数：
  extract(src, assets_dir, stem, keep_decor=False) -> str   # 返回 Markdown 文本
  epub_to_markdown(epub_path, assets_dir, stem, keep_decor=False) -> str
  azw3_to_markdown(azw3_path, assets_dir, stem, keep_decor=False) -> str

CLI（便于单独调试，不建议直接交付，仍应由 convert.py 走完整流水线）：
  python kindle_to_md.py "book.azw3" -o "book.md"
  python kindle_to_md.py "book.azw3" --keep-decor
"""
import os
import re
import sys
import shutil
import zipfile
import posixpath
import tempfile
import argparse
import warnings

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))

# calibre 转换器的 class 语义（见 style0001.css）
#   calibre1  = 居中块，多为章标题容器或纯图片块
#   calibre2  = 大号加粗（1.83em）→ 章标题
#   calibre3  = 首行缩进 2em 正文段
#   calibre4  = 101x101 小图（章首装饰）
#   calibre5  = 中号加粗（1.29em）→ 小节标题
CALIBRE_TITLE = {"calibre2"}
CALIBRE_SUBTITLE = {"calibre5"}
CALIBRE_DECOR_IMG = {"calibre4"}
CALIBRE_PAGEBREAK = {"mbp_pagebreak"}
# 章末整页彩图容器（推广页/二维码页）：内容即图片，无正文
FULLPAGE_IMG_CONTAINERS = {"zhangyue-c"}
# 注：`calibre1`（居中块）/ `calibre3`（首行缩进正文）**刻意不设常量**——
#   正文本就由 `<p>` 兜底分支处理，居中块除装标题外无额外语义。
#   2026-10-05 审计发现：这两个常量从未被任何代码读取（死常量），
#   却在操作指南的 class 映射表里被写成「本引擎处理」，属误导性文档。

SUPPORT_EXTS = {".azw3", ".mobi", ".kf8", ".epub"}


def _load(name):
    import importlib.util
    target = os.path.join(HERE, name)
    spec = importlib.util.spec_from_file_location(name, target)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _classes(el):
    try:
        return set(el.get("class") or [])
    except Exception:
        return set()


# ---------------------------------------------------------------- 资源落盘

class _Assets(object):
    """把书内图片复制到 assets_<基名>/，并给出相对路径。"""

    def __init__(self, assets_dir, stem, keep_decor=False):
        self.dir = assets_dir
        # Windows 不允许路径/文件名以空格或点结尾；书名常带「(精校本) 」这类尾随空格
        self.stem = re.sub(r"[ .]+$", "", stem) or "book"
        self.keep_decor = keep_decor
        self.rel = os.path.basename(assets_dir.rstrip("\\/"))
        self.n = 0
        self.saved = []
        os.makedirs(assets_dir, exist_ok=True)

    def save(self, abs_path, alt=""):
        if not abs_path or not os.path.isfile(abs_path):
            return None
        base = os.path.basename(abs_path)
        stem, ext = os.path.splitext(base)
        self.n += 1
        name = "%s_%02d_%s%s" % (self.stem, self.n, stem, ext)
        dst = os.path.join(self.dir, name)
        try:
            shutil.copyfile(abs_path, dst)
        except Exception as e:
            sys.stderr.write("[警告] 图片落盘失败 %s -> %s: %s\n" % (abs_path, dst, e))
            return None
        self.saved.append(dst)
        return "![%s](%s/%s)" % (alt or name, self.rel, name)


# ---------------------------------------------------------------- DOM -> Markdown

_BLOCKISH = {"p", "div", "section", "article", "body", "ul", "ol", "table",
             "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "hr"}


def _norm_text(el):
    """取元素纯文本：全角/半角空白与不换行空格统一压成单空格。"""
    t = el.get_text(" ", strip=True)
    return re.sub(r"[ \t\u00a0\u3000]+", " ", t).strip()


def _heading_child(el, target, max_len=0):
    """在 el **自身或直接子节点**中找第一个 class 命中 target 的元素，找不到返回 None。

    ★ 刻意**只查一层**（自身 + 直接子节点），不用 `el.find(class_=...)`、也不递归扫
      `descendants`。两个原因，都是2026-10-05 审计实证踩出来的：
      ① `find(class_=回调)` 在 beautifulsoup4 4.15.0 下对 `<p class=calibre1><span
         class=calibre2>标题</span></p>` 返回 None ⇒ 全部标题静默丢失，而门禁E0/W0 全绿；
      ② 若递归扫全部后代，则顶层 `<div>`（同时含标题与正文若干段）会命中标题后整块return，
         **同容器内的正文全部被吞掉**。
    max_len > 0 时额外要求该元素文本长度不超过它（用于小节标题的长护栏）。
    """
    cands = [el] + list(el.find_all(recursive=False))
    for c in cands:
        if not getattr(c, "name", None):
            continue
        if target & set(c.get("class") or []):
            # ★ 超长要`continue` 跳过该候选、继续看后面的候选，
            #   不能 `return None` —— 那会让「首个候选超长」直接终止整个查找，
            #   把后面本该命中的标题也丢掉（与本函数要消灭的「静默丢标题」同类问题）。
            if max_len and len(_norm_text(c)) > max_len:
                continue
            return c
    return None


def _img_md(im, assets, cur_dir):
    """<img> -> Markdown 图片语法；src 相对 cur_dir 解析，外链原样保留。"""
    src = im.get("src") or im.get("xlink:href") or ""
    if not src:
        return None
    if re.match(r"^[a-z]+:", src, re.I):     # http(s)/data 等
        return "![%s](%s)" % (im.get("alt") or "", src)
    abs_src = os.path.normpath(os.path.join(cur_dir, src.replace("/", os.sep)))
    return assets.save(abs_src, alt=im.get("alt") or "")


def _emit_heading(parent, child, assets, out, cur_dir, depth, level, front_matter=False):
    """把 child 渲染成标题（level=2/3），再把 parent 的其余子节点继续下钻。

    解决「顶层div 同时含标题 + 正文」时标题分支把同容器正文吞掉的问题（2026-10-05 审计）。
    """
    # 先出该子节点里的图（`--keep-decor` 时章首装饰图在这里）
    for im in child.find_all("img"):
        if assets.keep_decor:
            md = _img_md(im, assets, cur_dir)
            if md:
                out.append(md)
                out.append("")
    t = _norm_text(child)
    if t:
        out.append("#" * level + " " + t)
        out.append("")
    # 同容器其余子节点照常渲染
    for ch in parent.find_all(recursive=False):
        if ch is child:
            continue
        _render(ch, assets, out, cur_dir, depth + 1, front_matter)


def _render(el, assets, out, cur_dir, depth=0, front_matter=False):
    """递归渲染一个块级元素，产出 Markdown 行到 out。"""
    if depth > 40:
        return
    cls = _classes(el)
    tag = el.name

    # 分页标记：直接丢弃（不产出 md）
    if cls & CALIBRE_PAGEBREAK:
        return

    # 章末整页彩图容器（推广页/二维码页）：内容即图片，无正文
    if cls & FULLPAGE_IMG_CONTAINERS:
        for im in el.find_all("img"):
            if (_classes(im) & CALIBRE_DECOR_IMG) and not assets.keep_decor:
                continue
            md = _img_md(im, assets, cur_dir)
            if md:
                out.append(md)
                out.append("")
        return

    # ★ 章标题 / 小节标题：**只认自身与直接子节点**（见 `_heading_child` 的注释），
    #   命中后把该子节点渲染成标题，并**继续渲染 el 的其余子节点**——
    #   否则顶层容器同时含「标题 + 正文若干段」时，同容器正文会被整块吞掉。
    for child, level in ((_heading_child(el, CALIBRE_TITLE), 2),
                         (_heading_child(el, CALIBRE_SUBTITLE, 80), 3)):
        if child is None:
            continue
        if child is el:
            out.append("#" * level + " " + _norm_text(el))
            out.append("")
        else:
            _emit_heading(el, child, assets, out, cur_dir, depth, level, front_matter)
        return

    # 纯图片块（章首装饰小图 / 其它独立图）
    if not _norm_text(el) and el.find("img") is not None:
        for im in el.find_all("img"):
            if (_classes(im) & CALIBRE_DECOR_IMG) and not assets.keep_decor:
                continue
            md = _img_md(im, assets, cur_dir)
            if md:
                out.append(md)
                out.append("")
        return

    # 原生标题（<h1>~<h6>）
    if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
        t = _norm_text(el)
        if not t:
            return
        # 前置页（书名页/版权页/目录页）的 `<h1>` 是**页面标题**而非章节，
        #   保持 H1 会破坏「H1 唯一」，降级又会被误当一章 —— 故输出为普通段落。
        if front_matter:
            out.append(t)
            out.append("")
            return
        # ★★ 按「书内实际层级」归一，而不是无条件 +1（2026-10-05 二次修正）。
        #   问题背景：EPUB3 的原生标题层级**五花八门**——
        #     · 用 `<h1>` 表示章节（需降级，因书名已占 H1）；
        #     · 用 `<h2>` 表示章节、小节用 `<h3>`（**层级已正确**，再降级反而制造
        #       「H1 直接跳到 H3」的 `E-STR-02` 阻断 —— 实测这是**回归**）。
        #   规则：`_build` 先扫描全书记原生标题，取**最小层级** `base`，
        #   统一映射为「H2 起步」：`out_level = 2 + (raw_level - base)`。
        #     · 书用 h1 分章（base=1）⇒ h1→H2、h2→H3（与预期一致）；
        #     · 书用 h2 分章、小节 h3（base=2）⇒ h2→H2、h3→H3（**层级不动，无回归**）；
        #     · 深层级（h5 分章）⇒ h5→H2，层级被压平但**绝不跳级**。
        base = getattr(assets, "native_base_level", 1) or 1
        raw = int(tag[1])
        out_level = max(2, min(6, 2 + (raw - base)))
        out.append("#" * out_level + " " + t)
        out.append("")
        return

    if tag == "li":
        t = _norm_text(el)
        if t:
            out.append("- " + t)
            out.append("")
        return

    if tag == "hr":
        out.append("---")
        out.append("")
        return

    # 目录条目 / 目录标题
    if "sgc-toc-level" in cls:
        t = _norm_text(el)
        if t:
            out.append("- " + t)
            out.append("")
        return
    if "sgc-toc-title" in cls:
        t = _norm_text(el)
        if t:
            out.append("## " + t)
            out.append("")
        return

    # 正文段（calibre3 正文 / calibre1 居中段 / 其它 p）
    if tag == "p":
        t = _norm_text(el)
        if t:
            out.append(t)
            out.append("")
        return

    # 容器：若有块级子节点则下钻，否则整段当正文
    kids = list(el.find_all(recursive=False))
    if any(getattr(c, "name", "") in _BLOCKISH for c in kids):
        for ch in kids:
            _render(ch, assets, out, cur_dir, depth + 1, front_matter)
        return
    t = _norm_text(el)
    if t:
        out.append(t)
        out.append("")
        return
    # 空容器也可能裹着内联元素
    for ch in kids:
        _render(ch, assets, out, cur_dir, depth + 1, front_matter)


# 书籍前置页（书名页 / 版权页 / 目录页 / 封面）的识别。
#
# ★ 设计原则（2026-10-05 二次修正，审计实测驱动）：
#   只看**文件名**（basename），不看目录名；且**必须整体或按 `-`/`_` 分段精确匹配**。
#   原因：早期版本用 `re.search(整条相对路径)` + 无边界子串，实测大面积误伤 ——
#     · `cover` 无边界 ⇒ `dis-cover-y.xhtml`（discovery）被判成封面页，
#       该章标题**被永久丢弃**，而门禁 E0/W0/I0 **全绿零告警**（最严重的一类静默错误）；
#     · 目录名参与匹配 ⇒ `cover/Text/ch1.xhtml` 把**全部章节**误杀；
#     · `fm\d` 无边界 ⇒ `chapter_fm01.xhtml` 中间命中。
#   ⇒ 宁可**漏判**也**不误判**：漏判只多一个 H1 ⇒ 触发 `E-STR-01` **E 级阻断**（响亮失败、
#     用户立刻可见、可立即处理）；误判会**静默丢标题**而门禁全绿 E0/W0/I0（安静失败、最危险）。
# ★★ 前置页判定 =「归一化 → 段列表 → 首段查白名单」（2026-10-05 四轮迭代后定案）。
#
#   迭代史（每轮都被实测反例打回）：
#     ① `re.search(整条路径)` 无边界子串 ⇒ `dis-cover-y` / `cover/` 目录 / `chapter_fm01` 全误伤；
#     ② `^...$` 锚定 + 后缀放开 `[a-z]{1,4}` ⇒ `coverages` / `copyrighted` / `titlebar` 误伤；
#     ③ 单个正则拼接基础词+后缀 ⇒ 后缀正则与基础词互相抢匹配
#        （`nav-1` 不命中、`prefacexxx` 却命中），两种错误同时存在；
#     ④ 改「先sub 剥后缀再查词表」⇒ 同上。
#   ⇒ 最终形态：**不做正则拼接**，只做三步确定性操作 ——
#      1) 归一化分隔符（`_` / `.` / 空格 / 连续 `-` → 单个 `-`）；
#      2) 按 `-` 切成段；
#      3) **首段**（或首+尾段组合）查白名单。
#   这样「基础词」边界天然清晰：`nav` 命中、`navx`/`fmx`/`chapter` 不命中；
#   `nav-1` / `nav_ii` / `copyright_2024` 等带序数的写法由「首段命中 + 其余段是纯序数」覆盖。
_FRONT_MATTER_BASE = frozenset((
    # 目录 / 导航
    "nav", "toc", "contents", "ncx", "navigation", "tocncx",
    # 封面 / 书名页
    "cover", "coverpage", "titlepage", "halftitle", "title",
    # 版权 / 版本
    "copyright", "imprint", "colophon", "edition",
    # 致谢 / 声明类
    "dedication", "acknowledgement", "acknowledgment", "acknowledgements",
    "acknowledgments", "epigraph",
    # 前置部分
    "frontmatter", "front", "fm", "prelim", "prelims", "halves", "half",
    "foreword", "preface", "introduction", "prologue", "afterword",
))
# 「词 + page/front/matter」这类语义后缀
_FM_SEMANTIC_TAIL = frozenset((
    "page", "pages", "front", "matter", "nav", "toc", "cover", "title",
))
_FM_ORDINAL = re.compile(r"^(?:[0-9]{1,4}|[ivxlc]{2,5})$", re.I)
# 剥首段尾部的紧连序数（`nav1`→`nav`、`copyright2024`→`copyright`、`coverfront`→`coverfront` 不剥）
_FM_TRAILING_NUM = re.compile(r"[0-9]{1,4}$")
# 罗马数字作独立段时仍支持（`nav-ii` / `titlepage-i`），
# 但**紧连**在词后的罗马数字不剥离（`titlepageii` 会误判）


def _stem_is_front_matter(stem):
    """文件名（去扩展名、已转小写）是否属于书籍前置页。

    判定：归一化分隔符 → 取首段 → 首段命中白名单即True。
    其余段只允许「语义后缀」或「纯序数」，避免 `cover_story` / `copyrighted` 之类混入。
    """
    norm = re.sub(r"[-_.\s]+", "-", stem).strip("-").lower()
    if not norm:
        return False
    segs = [s for s in norm.split("-") if s]
    if not segs:
        return False
    head = segs[0]
    if head not in _FRONT_MATTER_BASE:
        # 允许「基础词 + 紧连序数」的写法（`nav1` / `fm2` / `copyright2024` / `coverfront`）：
        # 把首段尾部的数字/罗马数字剥掉再查一次词表。
        # ★ 必须**整段剥完再查**（贪婪到末尾），不能只剥后缀 ——
        #   否则 `copyrighted` 剥成 `copyright` + 残留 `ed` 会被误判。
        # ★ 只剥**阿拉伯数字**，不剥罗马数字：`titlepageii` / `copyrightiii` 剥罗马数字后会
        #   变成词表词而误判（实测多轮迭代后回退）。罗马数字写法按「宁可漏判」接受。
        stripped = _FM_TRAILING_NUM.sub("", head)
        if not stripped or stripped == head or stripped not in _FRONT_MATTER_BASE:
            return False
        head = stripped
    segs = [head] + segs[1:]
    # 首段已命中；其余段必须是语义后缀或纯序数，否则视为「碰巧以该词开头」的正常章节名
    for seg in segs[1:]:
        if seg in _FM_SEMANTIC_TAIL or _FM_ORDINAL.match(seg):
            continue
        return False
    return True


_FRONT_MATTER_DIRS = frozenset((
    "frontmatter", "front_matter", "front", "fm", "prelims", "halves",
))


def _is_front_matter(href, opf_base=None, root=None):
    """判断该 spine 项是否书籍前置页（书名页/版权页/目录页/封面）。

    为什么要单独识别：这些页的原生 `<h1>` 是**页面标题**（如「目录」「版权页」），
    不是章节标题；原样输出会让 H1 计数暴涨（实测某书4 个 ⇒ `E-STR-01` 阻断），
    而一律降级又会被误读成一章。

    判定（任一命中即返回 True）：
      1. 任一**路径目录名**属于 `frontmatter` / `fm` 等白名单（精确相等）；
      2. **文件名**（去扩展名后）精确等于白名单词，或匹配带 `-`/`_`/数字边界的模式。

    ★ 只用「精确相等 + 有边界的模式」，**不做子串搜索** ——
      否则 `discovery.xhtml` 会被 `cover` 误伤，导致章节标题被静默丢弃。
    """
    try:
        #★ 只取 basename 做判定；目录仅在**精确等于**白名单目录名时才算前置页
        norm = str(href).replace("\\", "/").rstrip("/")
        parts = [p for p in norm.split("/") if p and p != "."]
        if not parts:
            return False
        dirs, stem_file = parts[:-1], parts[-1]
        for d in dirs:
            if d.lower() in _FRONT_MATTER_DIRS:
                return True
        stem = stem_file.rsplit(".", 1)[0] if "." in stem_file else stem_file
        return _stem_is_front_matter(stem.lower())
    except Exception:
        return False


def _xhtml_to_md(xhtml_text, cur_dir, assets, front_matter=False):
    """单个 xhtml -> Markdown 行列表。cur_dir 为该文件所在目录（解析 img相对 src）。"""
    from bs4 import BeautifulSoup
    warnings.filterwarnings("ignore")
    # lxml 优先，但**必须有 html.parser 回退**：按 requirements.txt 装的环境若缺 lxml，
    # 硬编码 "lxml" 会抛 FeatureNotFound 直接崩（2026-10-05 审计发现）。
    soup = None
    for parser in ("lxml", "html.parser"):
        try:
            soup = BeautifulSoup(xhtml_text, parser)
            break
        except Exception:
            continue
    if soup is None:
        soup = BeautifulSoup(xhtml_text, "html.parser")
    body = soup.find("body")
    if body is None:
        return []
    out = []
    for ch in body.find_all(recursive=False):
        _render(ch, assets, out, cur_dir, 0, front_matter)
    return out


# ---------------------------------------------------------------- 目录 / spine 解析

_NS_RE = re.compile(r"\{[^}]+\}")


def _parse_container_xml(zf):
    """container.xml -> OPF 路径。"""
    try:
        data = zf.read("META-INF/container.xml").decode("utf-8", "replace")
    except KeyError:
        return "content.opf"
    m = re.search(r'full-path\s*=\s*"([^"]+)"', data)
    return m.group(1) if m else "content.opf"


def _parse_opf(data):
    """解析 content.opf 文本。

    返回 (items, spine_indexes, meta)：
      items       = [(id, media_type, href), ...]，按 manifest 出现顺序
      spine_indexes = spine 中itemref 的下标序列（**章节顺序以spine 为准**，
                      不能按文件名排序 —— part0001/part0010 字典序会乱）
      meta        = {dc:title, dc:creator, dc:publisher, dc:language, dc:date}
    """
    items, order = [], {}
    for m in re.finditer(r"<item\b[^>]*?/>", data):
        tag = m.group(0)
        i = re.search(r'\bid\s*=\s*"([^"]+)"', tag)
        h = re.search(r'\bhref\s*=\s*"([^"]+)"', tag)
        t = re.search(r'\bmedia-type\s*=\s*"([^"]+)"', tag)
        if not (i and h):
            continue
        order[i.group(1)] = len(order)
        items.append((i.group(1), t.group(1) if t else "", h.group(1)))
    spine = []
    for m in re.finditer(r"<itemref\b[^>]*?/>", data):
        r = re.search(r'\bidref\s*=\s*"([^"]+)"', m.group(0))
        if r and r.group(1) in order:
            spine.append(order[r.group(1)])
    meta = {}
    for key in ("title", "creator", "publisher", "language", "date"):
        m = re.search(r"<dc:%s[^>]*>(.*?)</dc:%s>" % (key, key), data, re.S)
        if m:
            meta[key] = re.sub(r"\{[^}]+\}", "", m.group(1)).strip()
    # ★ 返回 spine 的**原始顺序**，绝不 sorted()。
    #   spine 是「阅读顺序」的权威声明（EPUB 规范：manifest 排资源、spine 排阅读序，
    #   两者可以不一致，如非线性电子书/把插页排在正文中间）。
    #   2026-10-05 审计发现：本曾误写 sorted(spine)，等于把 spine 重排回 manifest 顺序、
    #   spine 完全失效（实测 spine `0001→0010→0003→0002` 被渲染成 `0001→0002→0010→0003`）。
    return items, spine, meta


# ---------------------------------------------------------------- 解包

def _unpack_epub(epub_path, workdir):
    """epub -> 解包目录；返回 OEBPS 根（content.opf 所在目录）。"""
    with zipfile.ZipFile(epub_path) as zf:
        zf.extractall(workdir)
        opf_path = _parse_container_xml(zf)
    opf_abs = os.path.join(workdir, opf_path.replace("/", os.sep))
    if os.path.isfile(opf_abs):
        return os.path.dirname(opf_abs)
    found = _find_opf(workdir)
    return os.path.dirname(found) if found else workdir


def _unpack_azw3(azw3_path, workdir):
    """azw3/mobi/kf8 -> 解包目录；优先 KF8(mobi8)，回退 MOBI7(mobi7)。"""
    try:
        import mobi
    except ImportError as e:
        raise RuntimeError("缺少依赖 mobi（解包 azw3/mobi 专用），请安装：pip install mobi") from e
    tempdir, _ = mobi.extract(azw3_path)
    cands = [
        os.path.join(tempdir, "mobi8", "OEBPS"),
        os.path.join(tempdir, "mobi8"),
        os.path.join(tempdir, "mobi7", "OEBPS"),
        os.path.join(tempdir, "mobi7"),
    ]
    for c in cands:
        if os.path.isdir(c):
            dst = os.path.join(workdir, "book")
            shutil.copytree(c, dst)
            return dst
    for dp, _dn, fn in os.walk(tempdir):     # MOBI7 无 OEBPS 层：按 *.opf 定位
        for f in fn:
            if f.lower().endswith(".opf"):
                dst = os.path.join(workdir, "book")
                shutil.copytree(dp, dst)
                return dst
    raise RuntimeError("azw3 解包失败：未找到 OPF（若为 DRM 加密本则无法处理）")


def _find_opf(root):
    for cand in ("content.opf", "OEBPS/content.opf", "META-INF/container.xml"):
        p = os.path.join(root, cand.replace("/", os.sep))
        if os.path.isfile(p) and cand.endswith(".opf"):
            return p
    for dp, _dn, fn in os.walk(root):
        for f in fn:
            if f.lower() == "content.opf":
                return os.path.join(dp, f)
    return None


# ---------------------------------------------------------------- 主转换

def _scan_native_base_level(todo, max_files=60):
    """扫描正文页的原生 `<h1>`~`<h6>`，返回**最小层级**（作为全书记法基准）。

    只读文件前`max_files` 个（避免大书扫全库拖慢），只统计**非前置页**。
    ★ 找不到任何原生标题时返回 1（等价于「按 h1 起算」，对无原生标题的书无影响）。
    """
    best = None
    for path, is_front in todo[:max_files]:
        if is_front:
            continue
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                head = f.read(200000)
        except Exception:
            continue
        for m in re.finditer(r"<h([1-6])\b", head, re.I):
            lv = int(m.group(1))
            if best is None or lv < best:
                best = lv
    return best or 1


def _build(root, assets_dir, stem, keep_decor=False):
    assets = _Assets(assets_dir, stem, keep_decor)
    opf = _find_opf(root)
    meta, ordered = {}, []
    if opf:
        items, spine, meta = _parse_opf(open(opf, encoding="utf-8", errors="replace").read())
        ordered = [items[i] for i in spine] or items
    base = os.path.dirname(opf) if opf else root

    title = meta.get("title") or stem
    chunks = ["# " + title, ""]
    info = []
    for k, label in (("creator", "作者"), ("publisher", "出版社"),
                     ("language", "语言"), ("date", "出版日期")):
        v = meta.get(k)
        if v and not re.match(r"^\d{4}-\d{2}-\d{2}T", v):
            info.append("%s：%s" % (label, v))
    if info:
        chunks.append("> " + "　|　".join(info))
        chunks.append("")

    # ---- ① 先确定「待渲染文件 + 是否前置页」，再据此算原生标题基准层级 ----
    todo = []          # [(abs_path, front_matter)]
    if ordered:
        for _item_id, mtype, href in ordered:
            if mtype and ("xhtml" not in mtype and "html" not in mtype):
                continue
            path = os.path.normpath(os.path.join(base, href.replace("/", os.sep)))
            if not os.path.isfile(path):
                continue
            todo.append((path, _is_front_matter(href, base, root)))
    if not todo:      # 兜底：无 OPF 或 spine 全失效时按文件名顺序遍历
        files = []
        for dp, _dn, fn in os.walk(root):
            for f in fn:
                if f.lower().endswith((".xhtml", ".html", ".htm")):
                    files.append(os.path.join(dp, f))
        for path in sorted(files):
            todo.append((path, _is_front_matter(path, base, root)))

    # ---- ② 扫描**正文页**的原生标题层级，取最小值作为基准 ----
    #   为什么要扫描：EPUB3 的层级习惯不统一 —— 有的用 h1 分章（需降级），
    #   有的用 h2 分章 + h3 小节（**本身已正确**，无条件降级会造成「跳级」回归）。
    #   只统计**非前置页**，否则「版权页的 h1」会把基准拉到 1。
    base_level = _scan_native_base_level(todo)
    assets.native_base_level = base_level

    used = 0
    for path, is_front in todo:
        used += 1
        chunks.extend(_xhtml_to_md(open(path, encoding="utf-8", errors="replace").read(),
                                   os.path.dirname(path), assets, front_matter=is_front))

    md = "\n".join(chunks)
    md = re.sub(r"\n{3,}", "\n\n", md).replace("\r\n", "\n").strip()
    if not md.endswith("\n"):
        md += "\n"
    sys.stderr.write("[kindle] 章节文件 %d 个，图片落盘 %d 张\n" % (used, len(assets.saved)))
    return md


def extract(src, assets_dir, stem, keep_decor=False):
    """按扩展名分派，返回 Markdown 文本（不做自整理，交由上游流水线）。"""
    ext = os.path.splitext(src)[1].lower()
    if ext not in SUPPORT_EXTS:
        raise RuntimeError("不支持的电子书扩展名：%s" % ext)
    workdir = tempfile.mkdtemp(prefix="k2md_")
    try:
        root = _unpack_epub(src, workdir) if ext == ".epub" else _unpack_azw3(src, workdir)
        return _build(root, assets_dir, stem, keep_decor)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def epub_to_markdown(epub_path, assets_dir, stem, keep_decor=False):
    return extract(epub_path, assets_dir, stem, keep_decor)


def azw3_to_markdown(azw3_path, assets_dir, stem, keep_decor=False):
    return extract(azw3_path, assets_dir, stem, keep_decor)


def main():
    ap = argparse.ArgumentParser(description="omni-md-toolkit Kindle 电子书抽取引擎")
    ap.add_argument("src", help=".azw3 / .mobi / .kf8 / .epub")
    ap.add_argument("-o", "--output", default=None, help="输出 .md（默认同名同目录）")
    ap.add_argument("--assets-dir", default=None)
    ap.add_argument("--keep-decor", action="store_true", help="保留章首装饰小图")
    a = ap.parse_args()
    out = a.output or (os.path.splitext(a.src)[0] + ".md")
    src_stem = os.path.splitext(os.path.basename(a.src))[0]
    stem = os.path.splitext(os.path.basename(out))[0] or src_stem
    # 目录名去尾随空格/点：Windows 不允许，否则后续建文件必失败（Errno 2）
    assets = a.assets_dir or os.path.join(
        os.path.dirname(os.path.abspath(out)),
        "assets_" + re.sub(r"[ .]+$", "", src_stem))
    md = extract(a.src, assets, stem, a.keep_decor)
    with open(out, "w", encoding="utf-8") as f:
        f.write(md)
    sys.stderr.write("[已写] %s (%d 字符)\n" % (out, len(md)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
