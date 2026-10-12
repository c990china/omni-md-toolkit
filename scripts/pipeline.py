# -*- coding: utf-8 -*-
r"""20260906_全格式转MD一键流水线（omni-md-toolkit pipeline）

工具包：omni-md-toolkit（scripts/pipeline.py）
用途：将指定文件夹下的所有格式文件批量转 MD 并自动自整理，兼顾单文件转换与整文件夹批量流。

流程：
  1. 批量多引擎转换（DOCX / PPTX / PDF / XLSX / XLS / CSV / HTML / MSG / TXT / 图片 /
     AZW3 / MOBI / KF8 / EPUB —— 电子书走本地 kindle_to_md 引擎，不经 markitdown）
  2. 自动自整理（md_reorganizer 最新修复版）—— 修正断行、标题层级、表格保护与修补
  3. 广告与水印清理（ad_clean）
  4. 质量门禁（md_lint）—— 输出 E/W 级规则体检报告，生成汇总流水线报告

==== 使用方式 ====

对应命令：
  python scripts/pipeline.py "<文件夹>"

常用选项：
  --skip-exists        跳过已存在 <基名>.md 的文件（增量模式）
  --reorganize-only    只整理已有 .md，不重新转换
  --lint-only          只跑 lint 报告，不转换不整理
  --proj-md            套用项目规范（去行首>、列表转1）2）3））
  --out <输出目录>     指定输出目录（默认：与输入文件同目录）
  --no-lint            跳过 lint 步骤

产物命名：**每个源文件只产出一个 `<基名>.md`**（已自整理、已体检），
不再产出 `<基名>_整理.md` 之类的副本。扫描件/图片**不做 OCR、不调网关**，
由当前模型看图写逐页稿后重跑本命令即自动合并。

自整理是**强制**的：若某个文件的版式自整理失败，本流水线会把那份**未整理**的
`<基名>.md` 一并摘除（改名搁置），宁可没有产物，也绝不留下未整理的裸稿。
"""

import os
import sys
import argparse
import subprocess
import json
from pathlib import Path
from datetime import datetime
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

TOOLKIT_DIR = Path(__file__).parent.resolve()

SUPPORTED_EXTS = {
    ".pdf", ".docx", ".doc", ".pptx", ".xlsx", ".xls", ".csv",
    ".html", ".htm", ".msg", ".epub", ".rtf", ".odt", ".txt",
    # Kindle 电子书（2026-10-05 新增）：走本地原生抽取引擎 kindle_to_md
    ".azw3", ".mobi", ".kf8",
    ".png", ".jpg", ".jpeg", ".webp"
}


def find_script(role: str) -> Path:
    role_pure = role[:-3] if role.endswith(".py") else role
    for p in TOOLKIT_DIR.iterdir():
        if p.suffix == ".py":
            if p.name.endswith(role_pure + ".py") or p.name.endswith(f"_{role_pure}.py") or p.stem == role_pure:
                return p
    direct = TOOLKIT_DIR / (role if role.endswith(".py") else f"{role}.py")
    if direct.exists():
        return direct
    raise FileNotFoundError(f"找不到脚本 {role}（在 {TOOLKIT_DIR}）")


# Windows 下禁止子进程弹控制台窗口（否则批处理会一开一关地闪窗、抢焦点）。
NO_WIN = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0


def run_py(script_name: str, *args) -> subprocess.CompletedProcess:
    script = find_script(script_name)
    cmd = [sys.executable, str(script)] + [str(a) for a in args]
    return subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", check=False,
                          creationflags=NO_WIN)


def collect_files(target_dir: Path, recursive: bool = False):
    pattern = "**/*" if recursive else "*"
    files = []
    for p in target_dir.glob(pattern):
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
            # .md 本就不在 SUPPORTED_EXTS 里，这里只需排掉流水线自己的报告文件。
            if not p.name.startswith("pipeline_report"):
                files.append(p)
    return sorted(files)


def discard_unorganized(md_path: Path):
    """自整理失败时，把**未整理**的产物从「成品名」上摘掉。

    优先删除；若删除被拦截（某些映射盘会拦 os.remove），则改名为 `.unorganized-丢弃`
    搁置，确保交付目录里不会出现一份看起来像成品的未整理 `<基名>.md`。
    """
    try:
        os.remove(md_path)
        return "removed"
    except OSError:
        pass
    try:
        os.replace(str(md_path), str(md_path) + ".unorganized-丢弃")
        return "renamed"
    except OSError as e:
        print(f"  [警告] 无法摘除未整理产物 {md_path.name}：{e}（请手动处理）")
        return "failed"


def process_file(file_path: Path, args, stats: Counter, reports: list):
    stem = file_path.stem
    ext = file_path.suffix.lower()
    out_dir = Path(args.out) if args.out else file_path.parent
    # 产物只有一个：与源文件同名的 <基名>.md（转换后立即原地自整理）
    md_path = out_dir / f"{stem}.md"

    print(f"\n{'='*60}\n[处理] {file_path.name} ({ext})")

    if args.skip_exists and md_path.exists():
        print(f"  [跳过] 已存在产物: {md_path.name}")
        stats["skipped"] += 1
        return

    # 1. 转换阶段
    if not args.reorganize_only and not args.lint_only:
        # --no-clean：自整理由本流水线统一做（见下方第 2 步），避免整理两次。
        conv_args = [str(file_path), "-o", str(md_path), "--no-clean", "--no-gate"]
        if hasattr(args, "pdf_engine") and args.pdf_engine:
            conv_args += ["--pdf-engine", args.pdf_engine]
        if getattr(args, "keep_decor", False):
            conv_args += ["--keep-decor"]
        res = run_py("convert.py", *conv_args)
        if res.returncode == 3:
            # 3 = 「待接手」：扫描件需要视觉能力，页面已渲染并给出作业单，**不是失败**。
            print("  [待接手] 扫描件需要视觉能力，作业单如下：")
            for _line in (res.stderr.strip() or res.stdout.strip()).splitlines():
                print("    " + _line)
            stats["need_ai"] += 1
            reports.append(f"[待接手] {file_path.name}: 按作业单逐页写 page_NN_ocr.md 后重跑本命令即自动合并")
            return
        if res.returncode != 0:
            print(f"  [转换失败] {res.stderr.strip() or res.stdout.strip()}")
            stats["convert_failed"] += 1
            reports.append(f"[失败] 转换失败: {file_path.name}")
            return
        stats["converted"] += 1
        print(f"  [转换成功] -> {md_path.name}")
    else:
        if not md_path.exists():
            print(f"  [警告] 找不到待处理 MD: {md_path.name}")
            return

    # 2. 自整理阶段（**原地更新同一个文件**，绝不另存 _整理.md）
    target_for_lint = md_path
    if not args.lint_only:
        reorg_args = [str(md_path), "-o", str(md_path)]
        if args.proj_md:
            reorg_args.append("--proj-md")
        res = run_py("md_reorganizer.py", *reorg_args)
        if res.returncode == 0:
            print(f"  [自整理成功] -> {md_path.name}")
            stats["reorganized"] += 1
        else:
            # 自整理失败 ⇒ 绝不留下未整理的裸稿：把产物从成品名上摘除后按失败计。
            print(f"  [自整理失败] {res.stderr.strip() or res.stdout.strip()}")
            how = discard_unorganized(md_path)
            print(f"  [自整理失败] 已摘除未整理的 {md_path.name}（{how}）；本文件不产出成品。")
            stats["reorganize_failed"] += 1
            if not args.reorganize_only and not args.lint_only:
                stats["converted"] -= 1      # 产物已被摘除，不应再计为「转换成功」
            reports.append(f"[失败] 自整理失败: {file_path.name} -> 已摘除未整理的 {md_path.name}，本文件无成品")
            return

    # 3. 质量体检与门禁
    if not args.no_lint:
        lint_args = [str(target_for_lint), "--json"]
        # ★ 电子书（azw3/mobi/kf8/epub）自 2026-10-05 起也纳入**字符级无损比对**
        #   （md_lint.source_text 已支持这 4 种格式）。此前白名单不含它们，
        #   导致**批量路径下电子书丢字完全不报警**（单文件 convert.py 路径却会报）——
        #   同一份产物因入口不同而门禁结论不同，是典型的「静默不检查」。
        if ext in (".docx", ".doc", ".pdf", ".txt",
                   ".azw3", ".mobi", ".kf8", ".epub"):
            lint_args += ["--source", str(file_path)]
        res = run_py("md_lint.py", *lint_args)
        try:
            data = json.loads(res.stdout) if res.stdout.strip() else None
            # md_lint --json 输出的是「报告对象数组」，每项含 summary{E,W,I}
            if isinstance(data, list):
                items = [d for d in data if isinstance(d, dict)]
            elif isinstance(data, dict):
                items = [data]
            else:
                raise ValueError("未知的 JSON 结构: %s" % type(data).__name__)
            e_cnt = sum(int((d.get("summary") or {}).get("E", 0)) for d in items)
            w_cnt = sum(int((d.get("summary") or {}).get("W", 0)) for d in items)
            i_cnt = sum(int((d.get("summary") or {}).get("I", 0)) for d in items)
            status_str = f"E:{e_cnt} W:{w_cnt} I:{i_cnt}"
            if e_cnt > 0:
                print(f"  [门禁拦截] 发现严重问题: {status_str}")
                stats["lint_errors"] += 1
            else:
                print(f"  [门禁通过] {status_str}")
                stats["lint_passed"] += 1
            reports.append(f"[{'FAIL' if e_cnt > 0 else 'PASS'}] {file_path.name} -> {status_str}")
        except Exception as e:
            # 门禁结果不可解析 ⇒ 绝不计入「通过」，避免假 PASS
            print(f"  [门禁不可信] md_lint --json 解析失败（rc={res.returncode}）: {e}")
            print(f"  [门禁不可信] 该文件结论不可信，请单独运行 md_lint.py 复核；已计为「待人工」")
            stats["lint_unknown"] += 1
            reports.append(f"[UNKNOWN] {file_path.name} -> 门禁结果不可解析，需人工复核")


def main():
    parser = argparse.ArgumentParser(description="omni-md-toolkit 批量流水线")
    parser.add_argument("folder", help="待处理的文件夹路径")
    parser.add_argument("--out", default=None, help="输出目录（默认同原文件同目录）")
    parser.add_argument("--skip-exists", action="store_true", help="增量模式：跳过已存在 <基名>.md 的文件")
    parser.add_argument("--reorganize-only", action="store_true", help="只整理已有 MD，不重新转换")
    parser.add_argument("--lint-only", action="store_true", help="只跑质量门禁报告")
    parser.add_argument("--proj-md", action="store_true", help="套用规范（编号转1）2）3），去行首>）")
    parser.add_argument("--no-lint", action="store_true", help="跳过质量门禁")
    parser.add_argument("--pdf-engine", choices=["extract_text", "markitdown"], default="extract_text", help="文本型 PDF 提取引擎（默认 extract_text，排版更优）")
    parser.add_argument("-r", "--recursive", action="store_true", help="递归子目录")
    parser.add_argument("--keep-decor", action="store_true",
                        help="电子书：保留章首装饰小图（默认剔除）")
    args = parser.parse_args()

    target_dir = Path(args.folder).resolve()
    if not target_dir.exists() or not target_dir.is_dir():
        sys.stderr.write(f"[错误] 目录不存在: {args.folder}\n")
        sys.exit(1)

    files = collect_files(target_dir, args.recursive)
    if not files:
        print(f"[提示] 在 {target_dir} 下未发现支持的待处理文档。")
        sys.exit(0)

    print(f"发现 {len(files)} 个待处理文件。开始执行流水线...")
    stats = Counter()
    reports = []
    start_time = datetime.now()

    for f in files:
        process_file(f, args, stats, reports)

    duration = (datetime.now() - start_time).total_seconds()
    report_file = (Path(args.out) if args.out else target_dir) / "pipeline_report.txt"
    report_content = [
        f"全格式转MD批量流水线报告 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"目标目录: {target_dir}",
        f"耗时: {duration:.2f} 秒",
        f"文件总数: {len(files)} | 转换成功: {stats['converted']} | 转换失败: {stats['convert_failed']} | 待接手: {stats['need_ai']} | 跳过: {stats['skipped']}",
        f"自整理完成: {stats['reorganized']} | 自整理失败: {stats['reorganize_failed']} | 门禁通过: {stats['lint_passed']} | 门禁拦截(E级): {stats['lint_errors']} | 门禁不可信: {stats['lint_unknown']}",
        "-" * 60,
        "详细清单:"
    ] + reports

    with open(report_file, "w", encoding="utf-8") as rf:
        rf.write("\n".join(report_content) + "\n")

    print(f"\n{'='*60}\n[流水线完成] 耗时 {duration:.2f}s，汇总报告已写入: {report_file.name}")

    # 退出码：有真失败才非 0；「待接手」不算失败（扫描件需要视觉能力，不是错误）。
    if (stats["convert_failed"] or stats["reorganize_failed"]
            or stats["lint_errors"] or stats["lint_unknown"]):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
