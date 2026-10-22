#!/usr/bin/env python3
"""PPTX 三层提取：① 文本（按坐标排序）② 表格/SmartArt/图表 ③ 图片清单。
用法: python3 extract_pptx.py <file.pptx> [--from N] [--to M] [--images-dir DIR]
"""
import sys, zipfile, re, os, argparse
import xml.etree.ElementTree as ET

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A = NS['a']
P = NS['p']


def slide_key(name):
    m = re.search(r'slide(\d+)\.xml$', name)
    return int(m.group(1)) if m else 0


def txt_of(el):
    return ''.join(t.text or '' for t in el.iter(f'{{{A}}}t'))


def shape_anchor(sp):
    xfrm = sp.find(f'.//{{{A}}}off')
    if xfrm is None:
        return (10**9, 10**9)
    return (int(xfrm.get('y', 0)), int(xfrm.get('x', 0)))


def table_text(tbl):
    out = []
    for tr in tbl.findall(f'{{{A}}}tr'):
        cells = []
        for tc in tr.findall(f'{{{A}}}tc'):
            # 注意：<a:p> 嵌在 <a:txBody> 内，必须用 iter 而非 findall
            cells.append(' '.join(txt_of(p).strip()
                                  for p in tc.iter(f'{{{A}}}p')).strip())
        if any(cells):
            out.append(' | '.join(cells))
    return out


def parse_slide(xml_bytes):
    root = ET.fromstring(xml_bytes)
    spTree = root.find(f'.//{{{P}}}cSld/{{{P}}}spTree')
    items = []
    if spTree is None:
        return items
    for node in list(spTree):
        tag = node.tag.split('}')[1]
        if tag == 'sp':
            t = txt_of(node).strip()
            if t:
                items.append((shape_anchor(node), 'text', t))
        elif tag in ('graphicFrame', 'pic', 'grpSp'):
            tables = node.findall(f'.//{{{A}}}tbl')
            for tbl in tables:
                rows = table_text(tbl)
                if rows:
                    items.append((shape_anchor(node), 'table', rows))
            if not tables:
                pics = node.findall(f'.//{{{P}}}pic') + node.findall(f'.//{{{A}}}blip')
                if pics:
                    items.append((shape_anchor(node), 'pic', f'[图片 x{len(pics)}]'))
                else:
                    t = txt_of(node).strip()
                    if t:
                        items.append((shape_anchor(node), 'text', t))
    items.sort(key=lambda i: i[0])
    return items


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--from', dest='frm', type=int, default=1)
    ap.add_argument('--to', dest='to', type=int, default=10**6)
    ap.add_argument('--images-dir', default=None)
    args = ap.parse_args()

    z = zipfile.ZipFile(args.pptx)
    names = z.namelist()
    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                    key=slide_key)

    # 附属结构：SmartArt / 图表 / 备注
    extras = {
        'diagrams': [n for n in names if n.startswith('ppt/diagrams/data')],
        'charts': [n for n in names if re.match(r'ppt/charts/chart\d+\.xml$', n)],
        'notes': [n for n in names if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n)],
        'media': [n for n in names if n.startswith('ppt/media/')],
    }
    print(f'### FILE: {os.path.basename(args.pptx)}')
    print(f'### slides={len(slides)} extras: '
          + ', '.join(f'{k}={len(v)}' for k, v in extras.items()))
    print()

    if args.images_dir:
        os.makedirs(args.images_dir, exist_ok=True)

    for s in slides:
        n = slide_key(s)
        if n < args.frm or n > args.to:
            continue
        print(f'===== SLIDE {n} =====')
        for _, kind, payload in parse_slide(z.read(s)):
            if kind == 'table':
                print('  [TABLE]')
                for row in payload:
                    print('   | ' + row)
            elif kind == 'pic':
                print(f'  {payload}')
            else:
                for line in payload.split('\n'):
                    if line.strip():
                        print('  ' + line.strip())
        # 备注
        notes_name = f'ppt/notesSlides/notesSlide{n}.xml'
        if notes_name in names:
            nt = txt_of(ET.fromstring(z.read(notes_name))).strip()
            nt = re.sub(r'\s*\d+\s*$', '', nt).strip()
            if nt and not nt.isdigit():
                print(f'  [NOTES] {nt}')
        print()

    # 图表数据
    for c in extras['charts']:
        root = ET.fromstring(z.read(c))
        vals = [t.text for t in root.iter(f'{{{A.replace("main","chart")}}}pt')]
        cats = [t.text for t in root.iter(
            '{http://schemas.openxmlformats.org/drawingml/2006/chart}pt')]
        if cats:
            print(f'### {c}: ' + ', '.join(x for x in cats if x)[:2000])

    if args.images_dir:
        for m in extras['media']:
            with open(os.path.join(args.images_dir, os.path.basename(m)), 'wb') as f:
                f.write(z.read(m))
        print(f'### media extracted -> {args.images_dir} ({len(extras["media"])} files)')


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == '__main__':
    main()
