# -*- coding: utf-8 -*-
"""md → 图文版 HTML（供 Chrome 打印成 PDF）。

把「开卷手册」这类长文档里**关系型**的内容渲染成图：
- 章节标题下的 `FLOWS = {章节号: [(节点, 说明), ...]}` → 横向流程条（AT A GLANCE）
- `**标签** — 正文` 的段落 → 左侧竖线卡片
- `>` 引用 → 原文定义（短句竖线、长段方框）
- 表格 → 隔行底色、表头跨页重复、列宽自适应
其余段落、列表原样保留。

用法：
    python3 handbook_html.py 输入.md 输出.html [--flows flows.json]
可选 flows.json：{"3": [["节点","说明"], ...], ...}，章节号取自 `## 3. ...`。
"""
import sys, re, io, json, html

CSS = """
@page { size: A4; margin: 12mm 13mm; }
body { font-family: "Songti SC", serif; font-size: 9.6pt; line-height: 1.58; color:#1b1b1b; margin:0; }
h1 { font-size: 22pt; margin: 6mm 0 2mm; }
h2 { font-size: 14pt; margin: 7mm 0 2.6mm; padding-bottom:1.6mm;
     border-bottom:1.2pt solid #333; break-after: avoid; }
h2 + .fighdr { margin-top: 1mm; }
h3 { font-size: 10.5pt; margin: 3mm 0 1.2mm; }
.sec { margin: 0 0 6mm; }
.sec.brk { break-before: page; }
.sec.brk h2, .sec:first-child h2 { margin-top: 0; }
p { margin: 0 0 2mm; }
.sep { color:#888; }
/* 卡片 */
.kv { display:grid; grid-template-columns: 42mm 1fr; gap:2.4mm; break-inside: avoid;
      margin: 0 0 2.8mm; border-left: 2.2pt solid #999; padding-left: 3mm; }
.kv b { font-size:9.2pt; }
.kv p { margin:0; font-size:9.2pt; color:#333; }
/* 原文定义 */
.qline { border-left: 1.6pt solid #666; padding-left: 2.8mm; margin: 0 0 2.2mm;
         font-size: 8.7pt; break-inside: avoid; }
.model { border: .8pt solid #444; background:#f4f4f4; padding: 2.5mm 3mm;
         margin: 1.5mm 0 3mm; font-size: 8.9pt; break-inside: auto; }
/* 表格 */
.mx { width:100%; border-collapse: collapse; margin: 1mm 0 3mm; font-size: 8.6pt; break-inside: auto; }
.mx thead { display: table-header-group; }      /* 跨页重复表头 */
.mx tr { break-inside: avoid; }
.mx.keep { break-inside: avoid; }                /* ≤12 行整表不拆 */
.mx.compact { font-size: 8.2pt; }                /* >30 行用紧凑样式 */
.mx th, .mx td { border: .6pt solid #aaa; padding: 1.4mm 1.8mm; text-align:left; vertical-align: top; }
.mx th { background:#e9e9e9; font-weight:700; }
.mx tr:nth-child(even) td { background:#f7f7f7; }
/* 流程图 */
.fighdr { font-size:7.8pt; letter-spacing:.8px; color:#666; margin: 2mm 0 1.4mm; text-transform:uppercase; }
.flow-h { display:flex; gap:1.8mm; margin: 0 0 3.5mm; break-inside: avoid; }
.flow-h .node { flex:1; border:.8pt solid #333; background:#fafafa; padding:1.8mm 2mm; }
.node b { display:block; font-size:9.2pt; }
.node span { display:block; font-size:8.3pt; color:#444; margin-top:.6mm; }
.arw { align-self:center; color:#555; font-size:8pt; }
/* 目录 */
.toc { column-count:2; column-gap:8mm; font-size:9pt; }
.tocrow { display:flex; gap:2mm; border-bottom:.4pt dotted #bbb; padding:.8mm 0; break-inside:avoid; }
.tnum { width:7mm; text-align:right; color:#666; }
.tpg { margin-left:auto; color:#666; font-variant-numeric: tabular-nums; }
.alabel { break-after: avoid; font-size:9pt; font-weight:700; margin: 2mm 0 1mm; }
"""

def inline(t):
    t = html.escape(t)
    t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
    t = t.replace(' ｜ ', ' <span class="sep">｜</span> ')
    return re.sub(r'\*(.+?)\*', r'<em>\1</em>', t)

def parse(md):
    secs, cur, i, lines = [], None, 0, md.split('\n')
    while i < len(lines):
        ln = lines[i]
        if ln.startswith('## '):
            cur = {'title': ln[3:].strip(), 'blocks': []}; secs.append(cur); i += 1; continue
        if cur is None: i += 1; continue
        if not ln.strip(): i += 1; continue
        if ln.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].startswith('|'):
                if not re.match(r'^\|[\s\-:|]+\|$', lines[i]):
                    rows.append([c.strip() for c in lines[i].strip('|').split('|')])
                i += 1
            cur['blocks'].append(('table', rows)); continue
        if ln.startswith('>'):
            cur['blocks'].append(('quote', ln.lstrip('>').strip())); i += 1; continue
        if ln.startswith('- '):
            items = []
            while i < len(lines) and lines[i].startswith('- '):
                items.append(lines[i][2:].strip()); i += 1
            cur['blocks'].append(('list', items)); continue
        if ln.startswith('### '):
            cur['blocks'].append(('h3', ln[4:].strip())); i += 1; continue
        cur['blocks'].append(('p', ln.strip())); i += 1
    return secs

def flow_html(nodes):
    out = ['<div class="fighdr">AT A GLANCE ｜ 一图看懂</div><div class="flow-h">']
    for k, (h, b) in enumerate(nodes):
        out.append(f'<div class="node"><b>{inline(h)}</b><span>{inline(b)}</span></div>')
        if k < len(nodes) - 1: out.append('<div class="arw">▶</div>')
    return ''.join(out) + '</div>'

def table_html(rows):
    if not rows: return ''
    cls = 'mx keep' if len(rows) <= 12 else ('mx compact' if len(rows) > 30 else 'mx')
    out = [f'<table class="{cls}"><thead><tr>']
    out += [f'<th>{inline(c)}</th>' for c in rows[0]]
    out.append('</tr></thead><tbody>')
    for r in rows[1:]:
        out.append('<tr>' + ''.join(f'<td>{inline(c)}</td>' for c in r) + '</tr>')
    return ''.join(out) + '</tbody></table>'

def main():
    src, dst = sys.argv[1], sys.argv[2]
    flows = {}
    if '--flows' in sys.argv:
        flows = {int(k): v for k, v in json.load(io.open(sys.argv[sys.argv.index('--flows') + 1], encoding='utf-8')).items()}
    md = io.open(src, encoding='utf-8').read()
    secs = parse(md)
    out = ['<!doctype html><html lang="zh"><head><meta charset="utf-8"><style>', CSS, '</style></head><body>']
    for sec in secs:
        n = sec['title'].split('.')[0].strip()
        n = int(n) if n.isdigit() else None
        title = sec['title']
        if title.startswith('Contents'):
            out.append('<section class="sec"><h2>' + inline(title) + '</h2><div class="toc">')
            for s in secs:
                if s['title'].startswith('Contents'): continue
                num = s['title'].split('.')[0].strip()
                out.append(f'<div class="tocrow"><span class="tnum">{num}</span>'
                           f'<span class="tt">{inline(".".join(s["title"].split(".")[1:]).strip())}</span>'
                           f'<span class="tpg" data-sec="{num}">--</span></div>')
            out.append('</div></section>'); continue
        out.append(f'<section class="sec" id="s{n}"><h2>{inline(title)}</h2>')
        if n in flows: out.append(flow_html(flows[n]))
        blocks = sec['blocks']
        for bi, (kind, val) in enumerate(blocks):
            nxt = blocks[bi + 1] if bi + 1 < len(blocks) else (None, None)
            if kind == 'table': out.append(table_html(val))
            elif kind == 'quote':
                cls = 'model' if len(val) > 220 else 'qline'
                out.append(f'<div class="{cls}">{inline(val)}</div>')
            elif kind == 'list':
                out.append('<ul>' + ''.join(f'<li>{inline(x)}</li>' for x in val) + '</ul>')
            elif kind == 'h3': out.append(f'<h3>{inline(val)}</h3>')
            else:
                if nxt[0] == 'quote' and len(val) <= 70:
                    out.append(f'<p class="alabel">{inline(val)}</p>'); continue
                m = re.match(r'^\*\*(.+?)\*\*\s*—\s*(.+)$', val)
                out.append(f'<div class="kv"><b>{inline(m.group(1))}</b><p>{inline(m.group(2))}</p></div>'
                           if m else f'<p>{inline(val)}</p>')
        out.append('</section>')
    out.append('</body></html>')
    io.open(dst, 'w', encoding='utf-8').write('\n'.join(out))
    print('写入', dst)

if __name__ == '__main__':
    main()
