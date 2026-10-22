# -*- coding: utf-8 -*-
"""把 .docx 反向导出为 .md（保留标题层级、列表、表格）。

用途：工作区里有一批**只有 docx、没有 md 源**的「孤本」（如 <课程代码>_课程总结.docx、
课程总览.docx）。孤本的风险是：源信息更新后没人会去改它，它会悄悄过期。
本脚本把 docx 转回 md，使其成为可维护的源，之后再改 md → md_to_docx.py 生成 docx。

用法:
    python3 docx_to_md.py <input.docx> [output.md]
    python3 docx_to_md.py <input.docx> --stdout        # 只打印，不写文件

映射规则:
    Heading 1..6 / Title  →  #..######
    List Bullet / List Number → - / 1.
    表格 → GFM 管道表（首行作表头）
    其余段落按正文输出；空段落只保留一个，避免大片空行
"""
import os
import sys

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_block_items(parent):
    """按文档真实顺序交替产出 Paragraph 与 Table。"""
    from docx.oxml.ns import qn
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def cell_text(cell):
    """单元格内多段用 <br> 连接，避免破坏表格结构。"""
    parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    return "<br>".join(parts)


def table_to_md(t):
    rows = []
    for r in t.rows:
        rows.append([cell_text(c) for c in r.cells])
    if not rows:
        return []
    # 统一列数
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |",
           "| " + " | ".join(["---"] * width) + " |"]
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return out


def para_to_md(p):
    txt = p.text.strip()
    style = (p.style.name or "").lower()
    if not txt:
        return ""
    if style.startswith("title"):
        return f"# {txt}"
    if style.startswith("heading"):
        try:
            lvl = int(style.split()[-1])
        except ValueError:
            lvl = 1
        return "#" * min(max(lvl, 1), 6) + " " + txt
    if "list bullet" in style:
        return f"- {txt}"
    if "list number" in style:
        return f"1. {txt}"
    if "quote" in style:
        return f"> {txt}"
    return txt


def convert(path):
    doc = Document(path)
    lines, prev_blank = [], False
    for block in iter_block_items(doc):
        if isinstance(block, Table):
            lines += [""] + table_to_md(block) + [""]
            prev_blank = True
            continue
        md = para_to_md(block)
        if md == "":
            if not prev_blank:
                lines.append("")
                prev_blank = True
            continue
        lines.append(md)
        prev_blank = False
    # 收紧连续空行
    out, blank = [], False
    for l in lines:
        if l == "":
            if blank:
                continue
            blank = True
        else:
            blank = False
        out.append(l)
    return "\n".join(out).strip() + "\n"


def main():
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        sys.exit(1)
    to_stdout = "--stdout" in argv
    args = [a for a in argv if not a.startswith("--")]
    src = args[0]
    md = convert(src)
    if to_stdout:
        sys.stdout.write(md)
        return
    dst = args[1] if len(args) > 1 else os.path.splitext(src)[0] + ".md"
    with open(dst, "w", encoding="utf-8") as f:
        f.write(md)
    print(f"saved {dst}  ({len(md.splitlines())} 行)")


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == "__main__":
    main()
