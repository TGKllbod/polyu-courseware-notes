# -*- coding: utf-8 -*-
"""给 md_to_docx 生成的 docx 加排版控制（Word 侧）。

四件事：
1. 指定章节另起一页（`--break 1,14,15`，按 `## N.` 的编号）
2. 短段落整体不拆页（≤400 字），长段落允许跨页但保留孤行控制
3. 表格行不跨页断开；≤12 行的表整表不拆；大表首行为重复表头
4. 列宽按内容显示宽度分配（中文按 2 倍宽计）
5. 页脚写 `页码 / 总页数`（PAGE / NUMPAGES 域）

用法：
    python3 paginate_docx.py 文件.docx [--break 1,14,15]
"""
import sys, re
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def _el(tag):
    return OxmlElement(tag)

def add_footer(doc):
    for sec in doc.sections:
        f = sec.footer
        p = f.paragraphs[0] if f.paragraphs else f.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run()
        fld = _el('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE')
        r = _el('w:r'); t = _el('w:t'); t.text = '1'; r.append(t); fld.append(r)
        run._r.addnext(fld)
        tail = p.add_run(' / ')
        fld2 = _el('w:fldSimple'); fld2.set(qn('w:instr'), 'NUMPAGES')
        r2 = _el('w:r'); t2 = _el('w:t'); t2.text = '1'; r2.append(t2); fld2.append(r2)
        tail._r.addnext(fld2)

def auto_width(table, usable_cm=16.6):
    ncol = len(table.columns)
    if not ncol: return
    weights = []
    for ci in range(ncol):
        longest = 0
        for cell in table.columns[ci].cells:
            for line in cell.text.split('\n'):
                disp = sum(2 if '\u4e00' <= ch <= '\u9fff' else 1 for ch in line)
                longest = max(longest, disp)
        weights.append(min(max(longest, 6), 70))
    total = sum(weights) or 1
    widths = [max(2.0, usable_cm * w / total) for w in weights]
    scale = usable_cm / sum(widths)
    table.autofit = False
    for ci, w in enumerate(widths):
        w = w * scale
        table.columns[ci].width = Cm(w)
        for cell in table.columns[ci].cells:
            cell.width = Cm(w)

def main():
    path = sys.argv[1]
    breaks = set()
    if '--break' in sys.argv:
        breaks = {int(x) for x in sys.argv[sys.argv.index('--break') + 1].split(',') if x.strip().isdigit()}
    doc = Document(path)
    add_footer(doc)
    for t in doc.tables:
        auto_width(t)
    for p in doc.paragraphs:
        pf = p.paragraph_format
        if p.style.name == 'Heading 1':
            pf.keep_with_next = True
            num = p.text.strip().split('.')[0].strip()
            pf.page_break_before = num.isdigit() and int(num) in breaks
        else:
            txt = p.text.strip()
            pf.keep_together = len(txt) <= 400     # 长段落允许跨页
            pf.widow_control = True
            if 0 < len(txt) <= 60:
                pf.keep_with_next = True           # 短标签与其后内容同页
    for t in doc.tables:
        big = len(t.rows) > 12
        for ri, row in enumerate(t.rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn('w:cantSplit')) is None: trPr.append(_el('w:cantSplit'))
            if ri == 0 and trPr.find(qn('w:tblHeader')) is None: trPr.append(_el('w:tblHeader'))
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_together = True
                    if (not big) and ri < len(t.rows) - 1:
                        p.paragraph_format.keep_with_next = True
    doc.save(path)
    print('排版已应用：', path)

if __name__ == '__main__':
    main()
