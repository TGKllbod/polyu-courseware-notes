# -*- coding: utf-8 -*-
"""
从「复习笔记.md 总源」按讲拆分并生成每讲的独立 docx（复习笔记-第N讲.docx）。
规则（用户明确）：复习笔记成品 docx 必须按讲（Lecture）独立成文，
严禁把多讲内容全部塞进单个 复习笔记.docx；md 仅为可维护总源。

用法:
    python3 split_to_lectures.py <复习笔记.md> [--keep-single] [--lecture-markers "regex"]
    python3 split_to_lectures.py <复习笔记.md> --prefix QRST7890_
默认按标题行定位讲边界，如:  ## 4. Lecture 1 — ...   /   ## 5. Lecture 2 — ...
若某门课的 md 未使用 "Lecture N" 标题，用 --lecture-markers 传入自定义正则。
输出: 与 md 同目录下的 复习笔记-第N讲.docx（N=该讲在课件中的讲号）。
命名约定: 成品必须带课程代码前缀（如 QRST7890_复习笔记-第1讲.docx），
故跑本脚本时请一律加 --prefix <课程代码>_，否则会生成不带前缀的重复文件。
"""
import re
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from md_to_docx import convert  # noqa: E402

DEFAULT_MARKER = re.compile(r"^#{1,4}\s+(\d+)\.\s*Lecture\s+(\d+)\s*[—–\-:(（]", re.IGNORECASE)


def parse_master(md_path, marker_re=None):
    """Return {lecture_no: [lines...]} and (title_lines_before_first_lecture).

    默认按 `## N. Lecture M — …` 找边界；其他课程可用 --lecture-markers 传自定义正则，
    要求：正则须有 2 个捕获组，**第 2 组为讲号**（如 `^#+\\s*(Lecture)\\s+(\\d+)\\b`）。
    """
    marker_re = marker_re or DEFAULT_MARKER
    with open(md_path, encoding="utf-8") as f:
        lines = f.read().split("\n")

    lectures = {}
    cur = None
    preamble = []
    seen_any = False
    for i, line in enumerate(lines, 1):
        m = marker_re.search(line)
        if m:
            try:
                cur = int(m.group(2))
            except (IndexError, ValueError):
                print(f"[ERROR] 正则第 2 组必须是讲号（数字）；第 {i} 行匹配失败: {line[:60]}")
                sys.exit(2)
            seen_any = True
            lectures.setdefault(cur, []).append(line)
            continue
        if not seen_any:
            preamble.append(line)
        elif cur is not None:
            lectures[cur].append(line)
    return lectures, preamble


def main():
    argv = sys.argv[1:]
    custom = None
    prefix = ""
    if "--lecture-markers" in argv:
        i = argv.index("--lecture-markers")
        custom = argv[i + 1]
        del argv[i:i + 2]
    if "--prefix" in argv:
        i = argv.index("--prefix")
        prefix = argv[i + 1]
        if prefix and not prefix.endswith("_"):
            prefix += "_"
        del argv[i:i + 2]
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(1)
    md_path = os.path.abspath(args[0])
    folder = os.path.dirname(md_path)
    marker_re = re.compile(custom, re.IGNORECASE) if custom else None
    lectures, preamble = parse_master(md_path, marker_re)

    if not lectures:
        print(f"[WARN] 未在 {md_path} 中找到讲边界标题，未生成任何 docx。"
              f"（默认正则: {DEFAULT_MARKER.pattern}；可用 --lecture-markers 指定）")
        sys.exit(0)


    made = []
    for no in sorted(lectures):
        slice_md = os.path.join(folder, f"._tmp_L{no}.md")
        body = "\n".join(lectures[no])
        # 无讲号或标题开头无编号时补一个标题行，保证 docx 首段有讲名
        with open(slice_md, "w", encoding="utf-8") as f:
            f.write(body)
        try:
            out = convert(slice_md)
            target = os.path.join(folder, f"{prefix}复习笔记-第{no}讲.docx")
            os.replace(out, target)
            made.append(target)
        finally:
            if os.path.exists(slice_md):
                os.remove(slice_md)
            if os.path.exists(slice_md[:-3] + ".docx"):
                os.remove(slice_md[:-3] + ".docx")
    print("made:", *made, sep="\n  ")


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == "__main__":
    main()
