# -*- coding: utf-8 -*-
"""把 Markdown 转成 .docx（正文 Calibri + 中文字体微软雅黑；支持标题/表格/列表/粗体/引用）。

用法:
    python3 md_to_docx.py <in.md> [out.docx]     # 单文件转换（默认同名 .docx）
    python3 md_to_docx.py --all-notes            # ⚠️ 危险：把全库每门课的 复习笔记.md 各转成一个"合并版 复习笔记.docx"
                                                 #    本库规则严禁合并版（见 SKILL.md §五），一般不要用；
                                                 #    正常按讲拆分请用 split_to_lectures.py。

历史教训：本脚本原先没有任何参数解析，直接跑就会把全部课程的 复习笔记.md
各转出一个合并版 复习笔记.docx（违反"一讲一个 docx"规则）。现已改为必须显式传参。
"""
import re, glob, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

BASE = os.environ.get("POLYU_HOME", os.path.expanduser("~/PolyU"))  # 用 POLYU_HOME 指向你的课程根目录
CN = "Microsoft YaHei"
DARK = RGBColor(0x1F, 0x38, 0x64)
GRAY = RGBColor(0x55, 0x55, 0x55)


def new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.2); sec.right_margin = Cm(2.2)
    n = doc.styles["Normal"]
    n.font.name = "Calibri"; n.font.size = Pt(10.5)
    n.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), CN)
    return doc


def set_cn(run, size=None, bold=None, color=None, italic=None):
    run.font.name = "Calibri"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), CN)
    if size is not None: run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color is not None: run.font.color.rgb = color
    if italic is not None: run.italic = italic
    return run


# 行内标记：顺序有意义——代码 > 粗斜 > 粗 > 斜 > 转义星号
CODE_RE     = re.compile(r'(`[^`]*`)')
BOLDITAL_RE = re.compile(r'\*\*\*(?=\S)(.+?)(?<=\S)\*\*\*')
BOLD_RE     = re.compile(r'\*\*(?=\S)(.+?)(?<=\S)\*\*')
ITAL_RE     = re.compile(r'(?<!\*)\*(?=\S)([^*\n]+?)(?<=\S)\*(?!\*)')
ESC_RE      = re.compile(r'\\\*')


def _emphasis(s, bold=False, italic=False, depth=0):
    """把一段**无代码**文字切成 (文字, bold, italic) 片段。"""
    if depth > 3 or not s:
        return [(s, bold, italic)] if s else []
    for rx, mk in ((BOLDITAL_RE, (True, True)), (BOLD_RE, (True, italic)), (ITAL_RE, (bold, True))):
        out, pos = [], 0
        for m in rx.finditer(s):
            if m.start() > pos:
                out += _emphasis(s[pos:m.start()], bold, italic, depth + 1)
            nb = mk[0] if mk[0] else bold
            ni = mk[1] if mk[1] else italic
            out += _emphasis(m.group(1), nb, ni, depth + 1)
            pos = m.end()
        if pos:
            if pos < len(s):
                out += _emphasis(s[pos:], bold, italic, depth + 1)
            return out
    return [(ESC_RE.sub('*', s), bold, italic)]


def add_runs(paragraph, text, base_size=10.5, base_bold=False, base_color=None, base_italic=False):
    r"""行内格式：`代码` > ***粗斜*** > **粗体** > *斜体* > \* 转义星号。

    ⚠️ 切分顺序有意义：**先按反引号切**，只在非代码片段里识别星号标记——否则
    代码里的星号（如 `` `rswt_*` ``、`` `cfps2022*.dta` ``）会与同段另一个星号
    配成一对、被误当斜体，把代码片段吃掉（发现）。
    斜体与 ***粗斜*** 为后来新增：此前不认，APA 参考文献的期刊名与卷号、
    以及笔记里的 *强调* 全部渲染成带星号的字面量。
    """
    def flush(pieces, is_code=False):
        for txt, b, i in pieces:
            if not txt:
                continue
            if is_code:
                r = paragraph.add_run(txt)
                r.font.name = 'Consolas'
                r.font.size = Pt(max(6.0, base_size - 0.5))
                r.font.color.rgb = RGBColor(0x7A, 0x3E, 0x00)
                if b or base_bold:
                    r.bold = True
                if i or base_italic:
                    r.italic = True
            else:
                set_cn(paragraph.add_run(txt), size=base_size,
                       bold=b or base_bold, color=base_color,
                       italic=i or base_italic)

    pos = 0
    for m in CODE_RE.finditer(text):
        if m.start() > pos:
            flush(_emphasis(text[pos:m.start()]))
        flush([(m.group(1)[1:-1], False, False)], is_code=True)
        pos = m.end()
    if pos < len(text):
        flush(_emphasis(text[pos:]))


def convert(md_path, out_path=None):
    with open(md_path, encoding='utf-8') as f:
        lines = f.read().split('\n')
    doc = new_doc()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        # 水平分割线（--- / *** / ___）不作为正文渲染
        if re.match(r'^\s*([-*_])\1{2,}\s*$', line):
            i += 1
            continue
        if line.strip().startswith('|'):
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                tbl_lines.append(lines[i].strip())
                i += 1
            # 解析表格
            headers = None; rows = []
            for j, l in enumerate(tbl_lines):
                cells = [c.strip() for c in l.strip().strip('|').split('|')]
                if j == 1 and all(re.match(r'^:?-{2,}:?$', c) for c in cells if c):
                    headers = [c.strip() for c in tbl_lines[0].strip().strip('|').split('|')]
                    continue
                if headers is not None:
                    rows.append(cells)
                elif j == 0:
                    continue
            if headers:
                ncol = len(headers)
                widths = [16.5 / ncol] * ncol
                t = doc.add_table(rows=1, cols=ncol)
                t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
                for j, h in enumerate(headers):
                    c = t.rows[0].cells[j]; c.text = ""
                    add_runs(c.paragraphs[0], h, base_size=10, base_bold=True)
                for row in rows:
                    cells = t.add_row().cells
                    for j in range(ncol):
                        val = row[j] if j < len(row) else ""
                        cells[j].text = ""
                        add_runs(cells[j].paragraphs[0], val, base_size=10)
                for r in t.rows:
                    for j, c in enumerate(r.cells):
                        c.width = Cm(widths[j])
                doc.add_paragraph()
            continue
        m = re.match(r'^(#{1,4})\s+(.*)', line)
        if m:
            level = len(m.group(1)); text = m.group(2).strip()
            if level == 1:
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_runs(p, text, base_size=16, base_bold=True, base_color=DARK)
            elif level == 2:
                p = doc.add_heading(level=1)
                add_runs(p, text, base_size=13, base_bold=True, base_color=DARK)
            elif level == 3:
                p = doc.add_heading(level=2)
                add_runs(p, text, base_size=11.5, base_bold=True, base_color=RGBColor(0x1F,0x6F,0x8B))
            else:
                p = doc.add_heading(level=3)
                add_runs(p, text, base_size=11, base_bold=True)
            i += 1
            continue
        if line.startswith('>'):
            text = line.lstrip('>').strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            add_runs(p, text, base_size=10, base_italic=True, base_color=GRAY)
            i += 1
            continue
        if line.startswith('- '):
            p = doc.add_paragraph(style="List Bullet")
            add_runs(p, line[2:].strip())
            i += 1
            continue
        if not line.strip():
            i += 1
            continue
        p = doc.add_paragraph()
        add_runs(p, line)
        i += 1
    out = out_path or (md_path[:-3] + ".docx")
    if out.lower().endswith(".md"):
        out = out[:-3] + ".docx"
    doc.save(out)
    return out


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == "__main__":
    import sys
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        sys.exit(1)
    if argv[0] == "--all-notes":
        files = sorted(glob.glob(f"{BASE}/**/复习笔记.md", recursive=True))
        print("⚠️  --all-notes 会生成合并版 docx，违反 SKILL §五「一讲一个 docx」规则；如确需请自行确认后再用。")
        for f in files:
            print("saved", convert(f))
        print("DONE", len(files))
        sys.exit(0)
    src = argv[0]
    if not os.path.exists(src):
        print(f"[错误] 找不到文件: {src}")
        sys.exit(1)
    print("saved", convert(src, argv[1] if len(argv) > 1 else None))
