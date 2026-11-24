# -*- coding: utf-8 -*-
"""PDF 侧三件套：目录页码回写 / 页脚盖章 / 60% 分页规则。

子命令：
    render   把 HTML 用 Chrome 打印成 PDF
    toc      按渲染结果回写目录页码（HTML 里 `<span class="tpg" data-sec="N">` 占位）
    stamp    每页页脚：左侧「本页所含章节」，右侧「页码 / 总页数」
    rule60   反复执行：章节标题落在页面 60% 以下 → 该章另起一页（需配合 handbook_html 的 --break 机制）

用法：
    python3 deck_pdf_pageno.py render in.html out.pdf
    python3 deck_pdf_pageno.py toc   in.html out.pdf source.md
    python3 deck_pdf_pageno.py stamp out.pdf source.md
    python3 deck_pdf_pageno.py rule60 in.html out.pdf source.md
"""
import sys, re, io, json, subprocess, pathlib
import pymupdf

CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'

def norm(t):
    return re.sub(r'[^\u4e00-\u9fffA-Za-z]', '', t)

def render(html, pdf):
    subprocess.run([CHROME, '--headless=new', '--disable-gpu', '--no-pdf-header-footer',
                    f'--print-to-pdf={pdf}', 'file://' + pathlib.Path(html).resolve().as_posix()],
                   check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print('渲染完成', pdf)

def _heads(md_path):
    md = io.open(md_path, encoding='utf-8').read()
    return re.findall(r'^## (\d+\. .+)$', md, flags=re.M)

def _locate(pdf, heads, skip_first_pages=1):
    doc = pymupdf.open(pdf)
    pages = [norm(doc[i].get_text()) for i in range(doc.page_count)]
    out = {}
    for h in heads:
        num = h.split('.')[0]
        k = norm(re.sub(r'^[\d\.\s]+', '', h))[:12]
        hit = next((i + 1 for i in range(skip_first_pages, len(pages)) if k in pages[i]), None)
        if hit is None:
            k2 = re.sub(r'[^\u4e00-\u9fff]', '', h)[:8]
            hit = next((i + 1 for i in range(skip_first_pages, len(pages)) if k2 and k2 in pages[i]), None)
        out[num] = hit
    return out, doc.page_count

def toc(html, pdf, md):
    """渲染后把真实页码写回 HTML，直到稳定（最多 5 轮）。"""
    heads = _heads(md)
    for _ in range(5):
        real, pages_n = _locate(pdf, heads)
        h = io.open(html, encoding='utf-8').read()
        cur = dict(re.findall(r'data-sec="(\d+)">([^<]*)<', h))
        if all(str(real.get(k)) == v for k, v in cur.items()):
            print('目录已一致，页数', pages_n); return
        h = re.sub(r'<span class="tpg" data-sec="(\d+)">[^<]*</span>',
                   lambda m: '<span class="tpg" data-sec="%s">%s</span>' % (m.group(1), real.get(m.group(1), '?')), h)
        io.open(html, 'w', encoding='utf-8').write(h)
        render(html, pdf)
    print('目录已回写')

def stamp(pdf, md):
    doc = pymupdf.open(pdf)
    heads = _heads(md)
    real, _ = _locate(pdf, heads)
    starts = {}
    for h in heads:
        num = h.split('.')[0]
        title = ''.join(ch for ch in h.replace('—', '-').replace('｜', '|') if 32 <= ord(ch) < 127)
        starts[num] = (real.get(num) or 1, re.sub(r'^\d+\.\s*', '', title))
    order = sorted(starts.items(), key=lambda kv: kv[1][0])
    total = doc.page_count
    for i in range(total):
        pno = i + 1
        here = [(n, st, t) for n, (st, t) in order if st == pno]
        before = [(n, st, t) for n, (st, t) in order if st < pno]
        parts = []
        if before:
            n, st, t = before[-1]; parts.append(f'§{n} {t[:30]}')
        for n, st, t in here:
            if not parts or not parts[0].startswith(f'§{n} '):
                parts.append(f'§{n} {t[:30]}')
        label = ' · '.join(parts)[:86] or 'Contents'
        page = doc[i]; w, h = page.rect.width, page.rect.height; y = h - 26
        page.draw_line(pymupdf.Point(36, y - 10), pymupdf.Point(w - 36, y - 10),
                       color=(0.6, 0.6, 0.6), width=0.4)
        page.insert_text(pymupdf.Point(36, y + 2), label, fontsize=7, fontname='helv', color=(0.35, 0.35, 0.35))
        s = f'{pno} / {total}'
        tw = pymupdf.get_text_length(s, fontname='helv', fontsize=7.5)
        page.insert_text(pymupdf.Point(w - 36 - tw, y + 2), s, fontsize=7.5, fontname='helv', color=(0.1, 0.1, 0.1))
    doc.saveIncr()
    print('页脚已加：', total, '页')

def rule60(html, pdf, md, forced_file='.force_breaks'):
    """新章节标题落在页面 60% 以下 → 记入 forced_file 并重排，迭代至稳定。"""
    heads = _heads(md)
    for it in range(5):
        real, _ = _locate(pdf, heads)
        doc = pymupdf.open(pdf)
        low = []
        for h in heads:
            num = h.split('.')[0]
            pg = real.get(num)
            if not pg or pg > doc.page_count: continue
            page = doc[pg - 1]
            k = norm(re.sub(r'^[\d\.\s]+', '', h))[:14]
            y = None
            for b in page.get_text('blocks'):
                if k and k in norm(b[4]):
                    y = b[1] / page.rect.height; break
            if y and y > 0.60: low.append(int(num))
        print(f'轮{it+1}: 低起点章节 {low}')
        if not low: break
        forced = set()
        try:
            forced = {int(x) for x in io.open(forced_file).read().split()}
        except Exception:
            pass
        forced |= set(low)
        io.open(forced_file, 'w').write(' '.join(sorted(map(str, forced))))
        print('  已记录强制换页：', sorted(forced), '—— 请用 handbook_html 重新生成 HTML 后重跑')

def main():
    if len(sys.argv) < 2:
        print(__doc__); return
    cmd = sys.argv[1]
    if cmd == 'render': render(sys.argv[2], sys.argv[3])
    elif cmd == 'toc': toc(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == 'stamp': stamp(sys.argv[2], sys.argv[3])
    elif cmd == 'rule60': rule60(sys.argv[2], sys.argv[3], sys.argv[4])
    else: print(__doc__)

if __name__ == '__main__':
    main()
