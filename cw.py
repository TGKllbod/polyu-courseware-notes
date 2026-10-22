#!/usr/bin/env python3
"""cw — courseware workflow CLI.

跨平台（macOS / Windows / Linux）的课程资料整理命令行工具。

子命令:
  init     建立工作区骨架（课程文件夹 + 模板）
  doctor   环境自检（依赖 / LibreOffice / OCR / 字体）
  ingest   归档新课件到对应课程文件夹
  notes    从 复习笔记.md 按讲拆分生成 docx
  check    跑同步检查与缩写检查
  render   用 LibreOffice 把 pptx/docx 转成 PDF（用于实际查看排版）
  ocr      图片文字识别（macOS 用 Vision，其他平台用 tesseract）

示例:
  cw init --root ~/PolyU
  cw doctor
  cw ingest ~/Downloads/Lecture3.pptx
  cw notes APSS5202 --lecture 7
  cw check
  cw render deck.pptx
"""
from __future__ import annotations

import argparse
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
HERE = Path(__file__).resolve().parent
SCRIPTS = HERE / "skills" / "polyu-courseware-notes" / "scripts"

# ---------------------------------------------------------------- 基础设施


def _utf8() -> None:
    """Windows 控制台默认 cp936，强制 UTF-8 输出，避免中文乱码/崩溃。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass


def ok(msg: str) -> None:
    print(f"[OK]   {msg}", flush=True)


def warn(msg: str) -> None:
    print(f"[!]    {msg}", flush=True)


def bad(msg: str) -> None:
    print(f"[x]    {msg}", flush=True)


def info(msg: str) -> None:
    print(f"       {msg}", flush=True)


def resolve_root(arg: str | None) -> Path:
    if arg:
        return Path(arg).expanduser().resolve()
    env = os.environ.get("POLYU_HOME")
    if env:
        return Path(env).expanduser().resolve()
    return Path.home() / "PolyU"


def find_soffice() -> str | None:
    """按平台找 LibreOffice 可执行文件。"""
    candidates: list[str] = []
    if IS_WIN:
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            candidates += [str(Path(base) / "LibreOffice" / "program" / "soffice.exe"),
                           str(Path(base) / "LibreOffice" / "program" / "soffice.com")]
    elif IS_MAC:
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    else:
        candidates += ["/usr/bin/soffice", "/usr/bin/libreoffice", "/usr/local/bin/soffice"]
    for c in candidates:
        if Path(c).exists():
            return c
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return None


def run_soffice(args: list[str], profile_dir: Path) -> subprocess.CompletedProcess:
    """调用 LibreOffice；隔离用户配置目录，否则无 GUI 环境会报 DeploymentException。"""
    exe = find_soffice()
    if not exe:
        raise RuntimeError(
            "找不到 LibreOffice。请安装后重试：\n"
            "  Windows: winget install TheDocumentFoundation.LibreOffice\n"
            "  macOS:   brew install --cask libreoffice  "
            "或从 libreoffice.org 下载"
        )
    profile_dir.mkdir(parents=True, exist_ok=True)
    profile_uri = profile_dir.resolve().as_uri()  # 跨平台正确的 file:/// 形式
    env = os.environ.copy()
    env["HOME"] = str(profile_dir.parent)          # macOS / Linux
    env["USERPROFILE"] = str(profile_dir.parent)   # Windows
    env["TMPDIR"] = str(profile_dir.parent)
    cmd = [exe, f"-env:UserInstallation={profile_uri}", "--headless", "--norestore"] + args
    return subprocess.run(cmd, env=env, capture_output=True, text=True)


def run_script(name: str, args: list[str], cwd: Path) -> int:
    """调用 skills 目录下的既有脚本，避免重复实现。"""
    script = SCRIPTS / name
    if not script.exists():
        bad(f"找不到脚本：{script}")
        return 2
    cmd = [sys.executable, str(script)] + args
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"   # 子进程打印中文必需（Windows 默认 cp1252/cp936）
    env["PYTHONUTF8"] = "1"
    return subprocess.run(cmd, cwd=str(cwd), env=env).returncode


# ---------------------------------------------------------------- init

README_TMPL = """# {code} {name}

> 课程资料工作区。本文件记录**课程级信息**（考核、DDL、出勤、教师、上课时间地点），
> 知识点请写在 `复习笔记.md`，不要写进本文件。

## 基本信息
- 课程代码：{code}
- 课程名称：{name}
- 上课时间/地点：
- 教师：

## 考核构成
| 项目 | 权重 | 截止／日期 | 备注 |
| --- | --- | --- | --- |
|  |  |  |  |

## 文件说明
- `复习笔记.md` — 全课程知识点总源（可维护）
- `{code}_复习笔记-第N讲.docx` — 按讲拆分的成品，一讲一个文件
- 课件原件按下载原名归档
"""


def cmd_init(a: argparse.Namespace) -> int:
    root = resolve_root(a.root)
    (root / ".render").mkdir(parents=True, exist_ok=True)
    ok(f"工作区：{root}")
    created = 0
    for spec in a.courses or []:
        code, _, name = spec.partition(":")
        code, name = code.strip(), (name or code).strip()
        d = root / f"{code}_{name}".replace(" ", "_")
        if d.exists():
            info(f"已存在，跳过：{d.name}")
            continue
        d.mkdir(parents=True)
        (d / "README.md").write_text(
            README_TMPL.format(code=code, name=name), encoding="utf-8"
        )
        created += 1
        ok(f"新建课程：{d.name}")
    if not created and not a.courses:
        warn("未指定 --courses；可加 --courses ABCD1234:课程名 再运行")
    info(f"下一步：把课件放进 {root}/<课程文件夹>/，然后 cw ingest <文件>")
    return 0


# ---------------------------------------------------------------- doctor

def cmd_doctor(a: argparse.Namespace) -> int:
    print(f"cw doctor — {platform.system()} {platform.release()} / Python {platform.python_version()}")
    print()
    fails = 0

    # 依赖
    for mod, pkg in [("docx", "python-docx"), ("pptx", "python-pptx"),
                     ("pdfplumber", "pdfplumber"), ("fitz", "pymupdf"),
                     ("PIL", "Pillow")]:
        try:
            __import__(mod)
            ok(f"依赖 {pkg}")
        except Exception:
            bad(f"依赖缺失 {pkg}（pip install {pkg}）")
            fails += 1

    # LibreOffice
    exe = find_soffice()
    if exe:
        ok(f"LibreOffice：{exe}")
    else:
        warn("未找到 LibreOffice——render 不可用（PPT 排版无法实际查看）")
        info("Windows: winget install TheDocumentFoundation.LibreOffice")

    # OCR
    ocr = None
    if IS_MAC and (SCRIPTS / "vision_ocr.swift").exists():
        ocr = "macOS Vision（需 swiftc 编译）"
    tesseract = shutil.which("tesseract")
    if tesseract:
        ocr = (ocr + " + " if ocr else "") + f"tesseract：{tesseract}"
    if ocr:
        ok(f"OCR：{ocr}")
    else:
        warn("未找到 OCR 后端——图片型课件无法读字")
        info("Windows: winget install UB-Mannheim.TesseractOCR（装后重开终端）")

    # 中文字体
    if IS_WIN:
        ok("中文字体：Microsoft YaHei（系统自带）")
    elif IS_MAC:
        warn("Windows 字体缺失，生成 docx 时中文回退到 PingFang SC（脚本已处理）")
    else:
        info("Linux：建议安装 fonts-noto-cjk")

    # git
    if shutil.which("git"):
        ok("git 可用")
    else:
        warn("未找到 git——版本管理不可用")

    print()
    if fails:
        bad(f"{fails} 项依赖缺失")
        return 1
    ok("环境就绪")
    return 0



def run_script_capture(name: str, args: list[str], cwd: Path) -> str:
    """调用脚本并捕获输出（强制 UTF-8，Windows 才不会解码失败）。"""
    script = SCRIPTS / name
    if not script.exists():
        return ""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        r = subprocess.run([sys.executable, str(script)] + args, cwd=str(cwd), env=env,
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
    except Exception as e:  # noqa: BLE001
        return f"[提取失败] {e}"
    return (r.stdout or "") + (r.stderr or "")


def extract_summary(dest: Path, outdir: Path) -> list[str]:
    """对归档后的课件做四层提取，结果写进 <课程>/.render/extract/<name>.txt。"""
    suffix = dest.suffix.lower()
    outdir.mkdir(parents=True, exist_ok=True)
    out_file = outdir / (dest.stem + ".txt")
    lines: list[str] = []

    if suffix in (".pptx", ".potx"):
        text = run_script_capture("extract_pptx.py", [str(dest)], dest.parent)
        out_file.write_text(text, encoding="utf-8")
        m = re.search(r"### slides=(\d+) extras:\s*diagrams=(\d+), charts=(\d+), notes=(\d+), media=(\d+)", text)
        tables = text.count("[TABLE]")
        if m:
            slides, dias, charts, notes, media = m.groups()
            lines.append(f"页数 {slides} · 表格 {tables} · 图表 {charts} · 备注 {notes} · 图片 {media}")
            if notes != "0" and int(notes) < int(slides) // 3:
                lines.append(f"提示：备注只有 {notes} 条（共 {slides} 页）——先看几条再决定要不要依赖它")
        else:
            lines.append(f"已提取，但未解析到统计行（表格 {tables}）")
    elif suffix == ".pdf":
        try:
            import pdfplumber  # noqa: PLC0415
        except ImportError:
            return ["[!] 未安装 pdfplumber，跳过 PDF 提取"]
        chunks: list[str] = []
        pages = tables = chars = 0
        with pdfplumber.open(dest) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                t = page.extract_text() or ""
                chars += len(t)
                tb = page.extract_tables() or []
                tables += len(tb)
                chunks.append(f"===== PAGE {i} =====\n{t}")
                for ti, table in enumerate(tb, 1):
                    chunks.append(f"[TABLE {ti}] " + " | ".join(
                        " / ".join((c or "").replace("\n", " ").strip() for c in row) for row in table))
                pages += 1
        out_file.write_text("\n".join(chunks), encoding="utf-8")
        lines.append(f"页数 {pages} · 表格 {tables} · 文本 {chars} 字符")
        if pages and chars / pages < 50:
            lines.append("[!] 平均每页字符很少——可能是图片型 PDF，需要 OCR（cw ocr）")
    elif suffix in (".docx",):
        # docx_to_md 默认写文件只打印状态，必须加 --stdout 才拿得到内容
        text = run_script_capture("docx_to_md.py", ["--stdout", str(dest)], dest.parent)
        out_file.write_text(text, encoding="utf-8")
        lines.append(f"已转 md（{len(text)} 字符）")
    else:
        return []

    lines.append(f"提取结果：{out_file}")
    return lines


# ---------------------------------------------------------------- ingest

# 课程代码两种常见形态：APSS5201（字母+4位数字）、HTI5T04（字母+数字+字母+数字）。
# 不用 \b——_ 是词字符，AB CD1234_Lecture5 这类会判成无边界。
CODE_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z]{2,4}\d[A-Z]\d{2,3}|[A-Z]{2,6}\d{3,4})(?![0-9])")


def _guess_course(path: Path, root: Path) -> Path | None:
    """从文件名或所在文件夹里找课程代码，匹配工作区内的课程文件夹。"""
    m = CODE_RE.search(path.stem)
    code = m.group(1) if m else None
    if not code and path.parent != root:
        m2 = CODE_RE.search(path.parent.name)
        code = m2.group(1) if m2 else None
    if not code:
        return None
    for d in root.iterdir():
        if d.is_dir() and d.name.upper().startswith(code):
            return d
    return None


def cmd_ingest(a: argparse.Namespace) -> int:
    root = resolve_root(a.root)
    if not root.exists():
        bad(f"工作区不存在：{root}（先跑 cw init）")
        return 1
    rc = 0
    for raw in a.files:
        src = Path(raw).expanduser().resolve()
        if not src.exists():
            bad(f"文件不存在：{src}")
            rc = 1
            continue
        dest_dir = Path(a.course).resolve() if a.course else _guess_course(src, root)
        if not dest_dir or not dest_dir.is_dir():
            bad(f"无法判断 {src.name} 属于哪门课——用 --course <文件夹路径> 指定")
            rc = 1
            continue
        dest = dest_dir / src.name
        if dest.exists():
            warn(f"已存在同名文件，跳过：{dest.name}")
            if dest.stat().st_size == src.stat().st_size:
                info("大小相同，可能是重复下载；先比 md5 再决定是否删除")
            continue
        if a.dry_run:
            info(f"[dry-run] {src.name} -> {dest_dir.name}/")
            continue
        shutil.copy2(src, dest)
        size = dest.stat().st_size / 1024 / 1024
        ok(f"{src.name} -> {dest_dir.name}/（{size:.1f} MB）")
        if a.extract:
            for line in extract_summary(dest, dest_dir / ".render" / "extract"):
                info(line)
        if a.move:
            src.unlink()
            info("已删除源文件（--move）")
    if a.extract:
        info("下一步：把提取结果交给 AI 助手，按 skill 并入 复习笔记.md")
    else:
        info("下一步：加 --extract 可顺便做四层提取，或让 AI 助手按 skill 处理")
    return rc


# ---------------------------------------------------------------- notes

def cmd_notes(a: argparse.Namespace) -> int:
    root = resolve_root(a.root)
    target = Path(a.course)
    if not target.is_absolute():
        target = root / a.course
    if not target.is_dir():
        bad(f"找不到课程文件夹：{target}")
        return 1
    md = target / "复习笔记.md"
    if not md.exists():
        bad(f"找不到 {md.name}（总源）")
        return 1
    code = target.name.split("_")[0]
    args = [str(md), "--prefix", f"{code}_"]
    if a.lecture:
        args += ["--only", str(a.lecture)] if "--only" in _script_help("split_to_lectures.py") else []
    rc = run_script("split_to_lectures.py", args, root)
    if rc == 0:
        ok(f"已生成 {code}_复习笔记-*.docx")
    return rc


def _script_help(name: str) -> str:
    try:
        return subprocess.run([sys.executable, str(SCRIPTS / name), "--help"],
                              capture_output=True, text=True).stdout
    except Exception:
        return ""


# ---------------------------------------------------------------- check

def cmd_check(a: argparse.Namespace) -> int:
    root = resolve_root(a.root)
    if not root.exists():
        bad(f"工作区不存在：{root}")
        return 1
    print("=== sync_check（md 比 docx 新）===", flush=True)
    rc1 = run_script("sync_check.py", (["--fix"] if a.fix else []), root)
    print()
    print("=== acronym_check（缩写首现是否给全称）===", flush=True)
    rc2 = run_script("acronym_check.py", ["."], root)
    print()
    if rc1 == 0 and rc2 == 0:
        ok("检查通过")
        return 0
    bad("检查未通过，按上面的提示逐项修")
    return 1


# ---------------------------------------------------------------- render

def cmd_render(a: argparse.Namespace) -> int:
    src = Path(a.file).expanduser().resolve()
    if not src.exists():
        bad(f"文件不存在：{src}")
        return 1
    out = Path(a.out).expanduser().resolve() if a.out else src.parent / ".render"
    out.mkdir(parents=True, exist_ok=True)
    profile = out / "lo_profile"
    try:
        proc = run_soffice(["--convert-to", "pdf", "--outdir", str(out), str(src)], profile)
    except RuntimeError as e:
        bad(str(e))
        return 1
    pdf = out / (src.stem + ".pdf")
    if pdf.exists():
        ok(f"已渲染：{pdf}")
        info("用 PDF 阅读器（或 cw 之外的截图工具）实际看一遍排版——坐标正确不等于好看")
        return 0
    bad("渲染失败")
    if proc.stderr:
        info(proc.stderr.strip().splitlines()[-1][:300])
    return 1


# ---------------------------------------------------------------- ocr

def cmd_ocr(a: argparse.Namespace) -> int:
    img = Path(a.image).expanduser().resolve()
    if not img.exists():
        bad(f"文件不存在：{img}")
        return 1
    if IS_MAC:
        binary = HERE / ".bin" / "vision_ocr"
        source = SCRIPTS / "vision_ocr.swift"
        if not binary.exists() and source.exists():
            warn("首次使用需编译 Vision OCR（约 10 秒）")
            binary.parent.mkdir(parents=True, exist_ok=True)
            r = subprocess.run(["swiftc", "-O", str(source), "-o", str(binary),
                                "-module-cache-path", "/tmp/mc"],
                               capture_output=True, text=True)
            if r.returncode != 0:
                bad("编译失败")
                info((r.stderr or "").strip().splitlines()[-1][:300] if r.stderr else "")
                return 1
        if binary.exists():
            return subprocess.run([str(binary), str(img)]).returncode
    tess = shutil.which("tesseract")
    if tess:
        r = subprocess.run([tess, str(img), "stdout", "-l", "chi_sim+eng"],
                           capture_output=True, text=True)
        print(r.stdout)
        return r.returncode
    bad("没有可用的 OCR 后端")
    info("macOS：自带 Vision（cw ocr 会自动编译）")
    info("Windows：winget install UB-Mannheim.TesseractOCR")
    return 1


# ---------------------------------------------------------------- main

def main(argv: list[str] | None = None) -> int:
    _utf8()
    p = argparse.ArgumentParser(
        prog="cw", description="课程资料整理 CLI（跨平台）",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    p.add_argument("-r", "--root", help="工作区根目录（默认 $POLYU_HOME 或 ~/PolyU）")
    sub = p.add_subparsers(dest="cmd", required=True)

    # 让 --root 既能写在子命令前，也能写在子命令后（cw init --root X）
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("-r", "--root", help="工作区根目录（默认 $POLYU_HOME 或 ~/PolyU）")

    s = sub.add_parser("init", parents=[common], help="建立工作区骨架")
    s.add_argument("--courses", nargs="*", metavar="CODE:NAME", help="如 ABCD1234:课程名")
    s.set_defaults(func=cmd_init)

    s = sub.add_parser("doctor", parents=[common], help="环境自检")
    s.set_defaults(func=cmd_doctor)

    s = sub.add_parser("ingest", parents=[common], help="归档课件到课程文件夹")
    s.add_argument("files", nargs="+", help="课件路径")
    s.add_argument("--course", help="指定目标课程文件夹（默认按文件名里的课程代码推断）")
    s.add_argument("--move", action="store_true", help="归档后删除源文件")
    s.add_argument("--dry-run", action="store_true", help="只演示，不复制")
    s.add_argument("--extract", action="store_true",
                   help="归档后立刻做四层提取（pptx/pdf/docx），结果存到 <课程>/.render/extract/")
    s.set_defaults(func=cmd_ingest)

    s = sub.add_parser("notes", parents=[common], help="按讲拆分生成 docx")
    s.add_argument("course", help="课程文件夹名（或相对根目录的路径）")
    s.add_argument("--lecture", type=int, help="只生成某一讲（若脚本支持）")
    s.set_defaults(func=cmd_notes)

    s = sub.add_parser("check", parents=[common], help="同步检查 + 缩写检查")
    s.add_argument("--fix", action="store_true", help="自动重新生成过期 docx")
    s.set_defaults(func=cmd_check)

    s = sub.add_parser("render", parents=[common], help="LibreOffice 转 PDF")
    s.add_argument("file", help="pptx / docx / odt")
    s.add_argument("--out", help="输出目录（默认源文件旁的 .render/）")
    s.set_defaults(func=cmd_render)

    s = sub.add_parser("ocr", parents=[common], help="图片文字识别")
    s.add_argument("image")
    s.set_defaults(func=cmd_ocr)

    a = p.parse_args(argv)
    try:
        return a.func(a)
    except KeyboardInterrupt:
        print()
        warn("已中断")
        return 130


if __name__ == "__main__":
    sys.exit(main())
