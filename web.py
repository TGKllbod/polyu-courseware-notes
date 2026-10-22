#!/usr/bin/env python3
"""cw serve — 本地 Web 界面（无第三方依赖，只用标准库）。

功能:
  · 列出工作区里各课程的课件数与按讲笔记状态（已生成 / 已过期）
  · 拖入课件 → 归档 + 四层提取，就地显示统计
  · 一键重新生成某门课的 docx
  · 一键跑同步检查 + 缩写检查

只监听 127.0.0.1，不对外暴露；不做鉴权（本机单人使用）。
"""
from __future__ import annotations

import json
import os
import re
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

CONSOLE = r"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>cw — 课程资料</title>
<style>
  :root{--fg:#1a1a1a;--muted:#6b6b6b;--line:#e3e3e3;--bg:#fafafa;--accent:#2f6f4f;--warn:#b06a00}
  *{box-sizing:border-box}
  body{margin:0;font:15px/1.55 -apple-system,"Segoe UI","Microsoft YaHei",system-ui,sans-serif;color:var(--fg);background:var(--bg)}
  header{padding:20px 28px;border-bottom:1px solid var(--line);background:#fff}
  h1{margin:0 0 4px;font-size:19px;font-weight:650}
  .sub{color:var(--muted);font-size:13px}
  main{padding:22px 28px;max-width:1080px}
  .card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-bottom:14px}
  .row{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .grow{flex:1}
  .name{font-weight:650}
  .tag{font-size:12px;padding:1px 7px;border-radius:999px;border:1px solid var(--line);color:var(--muted)}
  .tag.ok{color:var(--accent);border-color:#bfe0cd;background:#f2faf6}
  .tag.bad{color:var(--warn);border-color:#f0d9b0;background:#fff8ec}
  ul.lec{list-style:none;margin:10px 0 0;padding:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:6px}
  ul.lec li{display:flex;gap:8px;align-items:baseline;font-size:13.5px;padding:4px 6px;border-radius:6px}
  ul.lec li:nth-child(odd){background:#fbfbfb}
  ul.lec .t{color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  button{font:inherit;font-size:13px;padding:6px 12px;border-radius:7px;border:1px solid var(--line);background:#fff;cursor:pointer}
  button:hover{border-color:#c9c9c9;background:#fcfcfc}
  button.primary{background:var(--accent);color:#fff;border-color:var(--accent)}
  button.primary:hover{filter:brightness(1.06)}
  #drop{border:2px dashed var(--line);border-radius:10px;padding:26px;text-align:center;color:var(--muted);background:#fff}
  #drop.hot{border-color:var(--accent);background:#f4fbf7;color:var(--accent)}
  pre{background:#111;color:#e8e8e8;padding:12px 14px;border-radius:8px;font-size:12.5px;overflow:auto;max-height:300px;margin:10px 0 0}
  select{font:inherit;font-size:13px;padding:6px 8px;border-radius:7px;border:1px solid var(--line);background:#fff}
  .empty{color:var(--muted);padding:20px 0}
  .editor{display:none;background:#fff;border:1px solid var(--line);border-radius:10px;margin-bottom:14px}
  .editor.open{display:block}
  .editor .bar{display:flex;align-items:center;gap:10px;padding:11px 16px;border-bottom:1px solid var(--line)}
  .editor .bar strong{font-size:14px}
  .panes{display:grid;grid-template-columns:1fr 1fr;gap:0;min-height:420px}
  @media(max-width:820px){.panes{grid-template-columns:1fr}}
  .panes textarea{width:100%;height:420px;border:0;border-right:1px solid var(--line);padding:14px 16px;
    font:13px/1.6 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;resize:vertical;outline:none}
  .preview{padding:14px 18px;overflow:auto;max-height:420px;font-size:14px}
  .preview h3,.preview h4,.preview h5{margin:14px 0 6px;line-height:1.35}
  .preview h3{font-size:17px} .preview h4{font-size:15px} .preview h5{font-size:14px;color:#444}
  .preview p{margin:6px 0}
  .preview ul,.preview ol{margin:6px 0 6px 20px;padding:0}
  .preview li{margin:3px 0}
  .preview code{background:#f2f2f2;padding:1px 5px;border-radius:4px;font-size:12.5px}
  .preview blockquote{margin:8px 0;padding:6px 12px;border-left:3px solid var(--line);color:#555;background:#fbfbfb}
  .preview table{border-collapse:collapse;margin:10px 0;font-size:13px;width:100%}
  .preview th,.preview td{border:1px solid var(--line);padding:5px 9px;text-align:left;vertical-align:top}
  .preview th{background:#f7f7f7}
  .preview hr{border:0;border-top:1px solid var(--line);margin:14px 0}
</style></head><body>
<header>
  <h1>课程资料</h1>
  <div class="sub" id="root">加载中…</div>
</header>
<main>
  <div class="card">
    <div class="row">
      <strong>拖入课件</strong>
      <span class="grow"></span>
      <select id="course"></select>
      <button class="primary" id="uploadBtn">选择文件</button>
      <input type="file" id="pick" hidden multiple>
    </div>
    <div id="drop" style="margin-top:12px">把课件拖到这里（pptx / pdf / docx）</div>
    <pre id="uplog" hidden></pre>
  </div>
  <div class="editor" id="editor">
    <div class="bar">
      <strong id="edTitle">复习笔记.md</strong>
      <span class="tag" id="edMeta"></span>
      <span class="grow"></span>
      <span id="edMsg" class="sub"></span>
      <button class="primary" id="edSave">保存</button>
      <button id="edClose">关闭</button>
    </div>
    <div class="panes">
      <textarea id="edText" spellcheck="false"></textarea>
      <div class="preview" id="edPreview"></div>
    </div>
  </div>
  <div id="courses"></div>
</main>
<script>
const $ = s => document.querySelector(s);
let state = null;

async function load(){
  const r = await fetch('/api/state'); state = await r.json();
  $('#root').textContent = state.root;
  const sel = $('#course');
  sel.innerHTML = state.courses.map(c => `<option value="${c.dir}">${c.dir}</option>`).join('');
  const box = $('#courses');
  if(!state.courses.length){ box.innerHTML='<div class="empty">工作区里还没有课程文件夹。先跑 <code>cw init --courses "ABCD1234:课程名"</code></div>'; return; }
  box.innerHTML = state.courses.map(c => `
    <div class="card" data-dir="${c.dir}">
      <div class="row">
        <span class="name">${c.dir}</span>
        <span class="tag">${c.files} 个文件</span>
        <span class="tag ${c.notes_md?'ok':'bad'}">${c.notes_md?'有 复习笔记.md':'缺 复习笔记.md'}</span>
        <span class="grow"></span>
        <button data-act="edit" ${c.notes_md?'':'disabled'}>打开笔记</button>
        <button data-act="notes" ${c.notes_md?'':'disabled'}>生成 docx</button>
        <button data-act="check">检查</button>
      </div>
      ${c.lectures.length ? `<ul class="lec">${c.lectures.map(l=>`
        <li><span class="tag ${l.docx?(l.stale?'bad':'ok'):''}">${l.docx?(l.stale?'过期':'已生成'):'缺'}</span>
        <span>§${l.n} L${l.m}</span><span class="t">${l.title}</span></li>`).join('')}</ul>`
        : '<div class="empty">复习笔记.md 里还没有 <code>## N. Lecture M — 标题</code> 形式的讲次</div>'}
    </div>`).join('');
  box.querySelectorAll('button[data-act]').forEach(b=>b.onclick=()=>act(b));
}

async function act(btn){
  const dir = btn.closest('[data-dir]').dataset.dir;
  const kind = btn.dataset.act;
  if(kind === 'edit'){ return openNote(dir); }
  btn.disabled = true; const label = btn.textContent; btn.textContent = '处理中…';
  try{
    const r = await fetch(`/api/${kind}`, {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({course: dir})});
    const j = await r.json();
    alert((j.ok?'':'失败：\n') + (j.output||''));
    await load();
  } finally { btn.disabled=false; btn.textContent=label; }
}

async function upload(files){
  const course = $('#course').value;
  if(!course){ alert('先选课程'); return; }
  const log = $('#uplog'); log.hidden=false; log.textContent='';
  for(const f of files){
    log.textContent += `→ ${f.name}\n`;
    const r = await fetch('/api/upload', {method:'POST',
      headers:{'X-Filename': encodeURIComponent(f.name), 'X-Course': encodeURIComponent(course)},
      body: f});
    const j = await r.json();
    log.textContent += (j.ok ? (j.lines||[]).map(l=>'   '+l).join('\n') : '   失败：'+j.error) + '\n';
  }
  await load();
}

$('#uploadBtn').onclick = () => $('#pick').click();
$('#pick').onchange = e => { upload([...e.target.files]); e.target.value=''; };
const drop = $('#drop');
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev, e=>{e.preventDefault(); drop.classList.add('hot');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev, e=>{e.preventDefault(); drop.classList.remove('hot');}));
drop.addEventListener('drop', e => upload([...e.dataTransfer.files]));

// ---------- 笔记预览与编辑 ----------
let editing = null;

function renderMd(src){
  const esc = s => s.replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
  const inline = s => s
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>');
  const L = esc(src).split('\n'), out = [];
  let i = 0;
  const isRow = r => /^\s*\|.*\|\s*$/.test(r);
  const cells = r => r.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => inline(c.trim()));
  while(i < L.length){
    const ln = L[i];
    if(/^\s*$/.test(ln)){ i++; continue; }
    let m;
    if((m = ln.match(/^(#{1,6})\s+(.*)$/))){
      const h = Math.min(m[1].length + 2, 5);
      out.push(`<h${h}>${inline(m[2])}</h${h}>`); i++; continue;
    }
    if(/^\s*([-*_]\s*){3,}$/.test(ln)){ out.push('<hr>'); i++; continue; }
    if(/^\s*>/.test(ln)){
      const b = [];
      while(i < L.length && /^\s*>/.test(L[i])){ b.push(inline(L[i].replace(/^\s*>\s?/, ''))); i++; }
      out.push('<blockquote>' + b.join('<br>') + '</blockquote>'); continue;
    }
    if(isRow(ln) && i + 1 < L.length && /^\s*\|[\s:|-]+\|\s*$/.test(L[i+1])){
      const rows = [];
      while(i < L.length && isRow(L[i])){ rows.push(L[i]); i++; }
      const head = cells(rows[0]);
      const body = rows.slice(2).map(cells);
      out.push('<table><thead><tr>' + head.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>'
        + body.map(r => '<tr>' + r.map(c => `<td>${c}</td>`).join('') + '</tr>').join('')
        + '</tbody></table>');
      continue;
    }
    if(/^\s*[-*]\s+/.test(ln)){
      const it = [];
      while(i < L.length && /^\s*[-*]\s+/.test(L[i])){ it.push(inline(L[i].replace(/^\s*[-*]\s+/, ''))); i++; }
      out.push('<ul>' + it.map(x => `<li>${x}</li>`).join('') + '</ul>'); continue;
    }
    if(/^\s*\d+\.\s+/.test(ln)){
      const it = [];
      while(i < L.length && /^\s*\d+\.\s+/.test(L[i])){ it.push(inline(L[i].replace(/^\s*\d+\.\s+/, ''))); i++; }
      out.push('<ol>' + it.map(x => `<li>${x}</li>`).join('') + '</ol>'); continue;
    }
    const para = [];
    while(i < L.length && !/^\s*$/.test(L[i]) && !/^(#{1,6}\s|\s*[-*]\s|\s*\d+\.\s|\s*>|\s*\|)/.test(L[i])){
      para.push(inline(L[i])); i++;
    }
    out.push('<p>' + para.join('<br>') + '</p>');
  }
  return out.join('\n');
}

async function openNote(dir){
  editing = dir;
  const r = await fetch('/api/note?course=' + encodeURIComponent(dir));
  const j = await r.json();
  if(!j.ok){ alert('打不开：' + j.error); return; }
  $('#edTitle').textContent = dir + ' / 复习笔记.md';
  $('#edText').value = j.content;
  $('#edMeta').textContent = (j.content.length / 1000).toFixed(1) + 'k 字符';
  $('#edMsg').textContent = '';
  $('#edPreview').innerHTML = renderMd(j.content);
  $('#editor').classList.add('open');
  $('#editor').scrollIntoView({behavior:'smooth', block:'nearest'});
  $('#edText').focus();
}

$('#edText').addEventListener('input', e => {
  $('#edPreview').innerHTML = renderMd(e.target.value);
  $('#edMeta').textContent = (e.target.value.length / 1000).toFixed(1) + 'k 字符';
  $('#edMsg').textContent = '未保存';
});

async function saveNote(){
  if(!editing) return;
  const btn = $('#edSave'); btn.disabled = true; btn.textContent = '保存中…';
  try{
    const r = await fetch('/api/note', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({course: editing, content: $('#edText').value})});
    const j = await r.json();
    if(!j.ok){ alert('保存失败：' + j.error); return; }
    $('#edMsg').textContent = `已保存（${j.bytes} 字节）`
      + (j.stale && j.stale.length ? ` · 第 ${j.stale.join('、')} 讲 docx 已过期` : '');
    await load();
  } finally { btn.disabled = false; btn.textContent = '保存'; }
}

$('#edSave').onclick = saveNote;
$('#edClose').onclick = () => { $('#editor').classList.remove('open'); editing = null; };
document.addEventListener('keydown', e => {
  if((e.metaKey || e.ctrlKey) && e.key === 's'){ e.preventDefault(); saveNote(); }
});

load();
</script></body></html>
"""

HEAD_RE = re.compile(r"^##\s+(\d+)\.\s*Lecture\s+(\d+)\s*[—–\-:：（(]\s*(.*)$")


def scan(root: Path) -> dict:
    courses = []
    for d in sorted(root.iterdir()):
        if not d.is_dir() or d.name.startswith("."):
            continue
        md = d / "复习笔记.md"
        lectures = []
        if md.exists():
            for line in md.read_text(encoding="utf-8", errors="replace").splitlines():
                m = HEAD_RE.match(line.strip())
                if m:
                    n, lec, title = int(m.group(1)), int(m.group(2)), m.group(3).strip()
                    docx = d / f"{d.name.split('_')[0]}_复习笔记-第{lec}讲.docx"
                    if not docx.exists():  # 兼容任意前缀
                        cand = list(d.glob(f"*_复习笔记-第{lec}讲.docx"))
                        docx = cand[0] if cand else docx
                    exists = docx.exists()
                    stale = exists and md.stat().st_mtime > docx.stat().st_mtime
                    lectures.append({"n": n, "m": lec, "title": title[:60],
                                     "docx": exists, "stale": stale})
        files = [f for f in d.iterdir() if f.is_file() and not f.name.startswith(".")]
        courses.append({"dir": d.name, "files": len(files),
                        "notes_md": md.exists(), "lectures": lectures})
    return {"root": str(root), "courses": courses}


class Handler(BaseHTTPRequestHandler):
    root: Path
    server_version = "cw"
    protocol_version = "HTTP/1.1"

    # ---------- 工具
    def _send(self, code: int, body: bytes, ctype: str = "application/json; charset=utf-8") -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code: int = 200) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def log_message(self, fmt, *args):  # 安静一点
        pass

    def _course(self, name: str) -> Path | None:
        """只允许访问根目录下第一层的目录，防路径穿越。"""
        if not name or "/" in name or "\\" in name or name.startswith("."):
            return None
        d = (self.root / name).resolve()
        try:
            d.relative_to(self.root.resolve())
        except ValueError:
            return None
        return d if d.is_dir() else None

    # ---------- 路由
    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send(200, CONSOLE.encode("utf-8"), "text/html; charset=utf-8")
        elif self.path == "/api/state":
            self._json(scan(self.root))
        elif self.path.startswith("/api/note?"):
            from urllib.parse import parse_qs  # noqa: PLC0415
            q = parse_qs(self.path.split("?", 1)[1])
            d = self._course((q.get("course") or [""])[0])
            if not d:
                return self._json({"ok": False, "error": "课程无效"}, 400)
            md = d / "复习笔记.md"
            if not md.exists():
                return self._json({"ok": False, "error": "缺 复习笔记.md"}, 404)
            return self._json({"ok": True, "content": md.read_text(encoding="utf-8", errors="replace"),
                               "mtime": md.stat().st_mtime})
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self):
        import cw  # 复用 cw 里的跑脚本/提取逻辑

        if self.path == "/api/upload":
            name = Path(_unq(self.headers.get("X-Filename", ""))).name
            course = _unq(self.headers.get("X-Course", ""))
            d = self._course(course)
            if not d or not name:
                return self._json({"ok": False, "error": "课程或文件名无效"}, 400)
            length = int(self.headers.get("Content-Length", 0))
            if length <= 0 or length > 500 * 1024 * 1024:
                return self._json({"ok": False, "error": "文件为空或过大"}, 400)
            data = self.rfile.read(length)
            dest = d / name
            dest.write_bytes(data)
            lines = cw.extract_summary(dest, d / ".render" / "extract")
            return self._json({"ok": True, "lines": [f"{name} 已归档（{len(data)/1048576:.1f} MB）"] + lines})

        if self.path == "/api/note":
            n = int(self.headers.get("Content-Length", 0))
            if n > 8 * 1024 * 1024:
                return self._json({"ok": False, "error": "内容过大（上限 8 MB）"}, 413)
            payload = json.loads(self.rfile.read(n) or b"{}")
            d = self._course(payload.get("course", ""))
            if not d:
                return self._json({"ok": False, "error": "课程无效"}, 400)
            content = payload.get("content")
            if not isinstance(content, str):
                return self._json({"ok": False, "error": "内容必须是字符串"}, 400)
            md = d / "复习笔记.md"
            tmp = md.with_suffix(".md.tmp")
            tmp.write_text(content, encoding="utf-8")   # 先写临时文件再原子替换，避免写一半
            os.replace(tmp, md)
            # 保存后告诉前端：哪些讲的 docx 现在过期了
            by_dir = {c["dir"]: c for c in scan(self.root)["courses"]}
            stale = [x["m"] for x in by_dir.get(d.name, {"lectures": []})["lectures"] if x["stale"]]
            return self._json({"ok": True, "bytes": len(content.encode("utf-8")), "stale": stale})

        if self.path in ("/api/notes", "/api/check"):
            n = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(n) or b"{}")
            d = self._course(payload.get("course", ""))
            if not d:
                return self._json({"ok": False, "error": "课程无效"}, 400)
            md = d / "复习笔记.md"
            if self.path == "/api/notes":
                if not md.exists():
                    return self._json({"ok": False, "error": "缺 复习笔记.md"}, 400)
                code = d.name.split("_")[0]
                out = cw.run_script_capture("split_to_lectures.py", [str(md), "--prefix", f"{code}_"], self.root)
                return self._json({"ok": "Traceback" not in out, "output": out.strip() or "完成"})
            out = cw.run_script_capture("sync_check.py", [], self.root)
            out += "\n" + cw.run_script_capture("acronym_check.py", ["."], self.root)
            return self._json({"ok": True, "output": out.strip()})

        self._json({"error": "not found"}, 404)


def _unq(s: str) -> str:
    from urllib.parse import unquote  # noqa: PLC0415
    return unquote(s or "")


def serve(root: Path, host: str = "127.0.0.1", port: int = 8765) -> int:
    Handler.root = root
    httpd = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"cw serve — {root}")
    print(f"打开 {url}   （Ctrl+C 退出）", flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止", flush=True)
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    r = Path(sys.argv[1]).expanduser().resolve() if len(sys.argv) > 1 else Path.home() / "PolyU"
    sys.exit(serve(r))
