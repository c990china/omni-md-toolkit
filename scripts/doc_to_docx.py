# -*- coding: utf-8 -*-
"""旧版 .doc（Word 97-2003 二进制）-> .docx 转换（Windows + Word COM）。

所属：omni-md-toolkit（scripts/doc_to_docx.py）
在已安装 Microsoft Office 的 Windows 环境下，通过 Word COM 直接将 .doc 转换为 .docx，
使后续能够走 docx_merge_table 完成高保真合并单元格修复与 Markdown 转换。

用法：
    python *_doc_to_docx.py "<文件或目录>" [-o 输出目录] [--recursive]

Python API:
    from ... import convert_one, convert_doc
    dst, ok = convert_doc(src_doc_path, out_dir=None)
"""
import argparse
import glob
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def convert_one(word, src, out_dir):
    base = os.path.splitext(os.path.basename(src))[0]
    dst = os.path.join(out_dir, base + ".docx")
    doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
    try:
        doc.SaveAs2(os.path.abspath(dst), FileFormat=16)  # 16 = wdFormatDocumentDefault
    finally:
        doc.Close(False)
    ok = os.path.exists(dst) and os.path.getsize(dst) > 0
    return dst, ok


def convert_doc(src, out_dir=None):
    """便捷 Python API：将单个 .doc 文件转换为 .docx"""
    if not os.path.isfile(src):
        raise FileNotFoundError(f"文件不存在: {src}")
    target_dir = out_dir or os.path.dirname(os.path.abspath(src))
    os.makedirs(target_dir, exist_ok=True)

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError("缺少依赖 pywin32，无法使用 Word COM 转换 .doc 文件。请安装：pip install pywin32")

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        dst, ok = convert_one(word, src, target_dir)
        return dst, ok
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main(argv=None):
    ap = argparse.ArgumentParser(description="旧版 .doc -> .docx（Word COM）")
    ap.add_argument("input", help=".doc 文件或包含 .doc 的目录")
    ap.add_argument("-o", "--output", default=None, help="输出目录，默认与源文件同目录")
    ap.add_argument("-r", "--recursive", action="store_true", help="递归子目录")
    args = ap.parse_args(argv)

    if os.path.isfile(args.input):
        files = [args.input]
    elif os.path.isdir(args.input):
        pat = os.path.join(args.input, "**", "*.doc") if args.recursive \
            else os.path.join(args.input, "*.doc")
        files = sorted(glob.glob(pat, recursive=args.recursive))
        # 排除 .docx 误匹配
        files = [f for f in files if not f.lower().endswith(".docx")]
    else:
        print("输入不存在: " + args.input)
        return 2

    if not files:
        print("未找到 .doc 文件")
        return 0

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("缺少 pywin32，请执行: pip install pywin32")
        return 3

    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    ok_n = 0
    try:
        for src in files:
            out_dir = args.output or os.path.dirname(os.path.abspath(src))
            if out_dir and not os.path.isdir(out_dir):
                os.makedirs(out_dir, exist_ok=True)
            try:
                dst, ok = convert_one(word, src, out_dir)
                print("  [%s] %s -> %s" % ("OK" if ok else "空文件",
                                           os.path.basename(src), dst))
                ok_n += ok
            except Exception as e:
                print("  [FAIL] %s : %r" % (os.path.basename(src), e))
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    print("成功 %d / %d" % (ok_n, len(files)))
    return 0 if ok_n == len(files) else 1


if __name__ == "__main__":
    sys.exit(main())
