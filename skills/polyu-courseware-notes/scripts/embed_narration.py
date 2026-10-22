#!/usr/bin/env python3
"""把逐页录音嵌入 PPTX，做成 PowerPoint「录制幻灯片演示」的效果。

产出结构与 PowerPoint for Mac 的「录制幻灯片演示」一致：
  * ppt/media/<name>.m4a        —— 音频本体
  * ppt/slides/_rels/slideN.xml.rels
        · .../relationships/audio              (a:audioFile 用)
        · http://schemas.microsoft.com/office/2007/relationships/media  (p14:media 用)
        · .../relationships/image              (音频图标的海报图，默认透明)
  * ppt/slides/slideN.xml
        · <p:pic>  音频承载形状，p:cNvPr@id = spid
        · <p:transition advTm="毫秒"/>  播完自动翻页
        · <p:timing> ... <p:audio isNarration="1">  旁白节点
  * [Content_Types].xml  注册音频扩展名
  * docProps/app.xml     MMClips = 音频数量

音频时长优先用 macOS 自带 afinfo 读取，失败则回退到 MP4 mvhd 原子解析。

用法：
    python3 embed_narration.py <deck.pptx> <audio_dir> [out.pptx]
    python3 embed_narration.py --verify <deck.pptx>

audio_dir 里放 1.m4a / 2.m4a / ... ，数字即幻灯片页码。
"""
from __future__ import annotations

import os
import re
import struct
import subprocess
import sys
import zlib
import zipfile

AUDIO_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

AUDIO_EXTS = {".m4a": "audio/mp4", ".mp4": "audio/mp4", ".m4b": "audio/mp4",
              ".mp3": "audio/mpeg", ".wav": "audio/wav", ".aac": "audio/aac",
              ".aiff": "audio/aiff", ".caf": "audio/x-caf"}

# 音频图标位置：与组员的 PowerPoint 输出一致（20in x 11.25in 画布右下角）
ICON_X, ICON_Y, ICON_CX, ICON_CY = 17322800, 9321800, 812800, 812800


def _png_1x1_transparent() -> bytes:
    """手工拼一个 1x1 全透明 RGBA PNG（不依赖 Pillow）。"""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0)      # 1x1, 8bit, RGBA
    idat = zlib.compress(b"\x00\x00\x00\x00\x00")            # filter 0 + 全 0 像素
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", idat) + chunk(b"IEND", b""))


def audio_duration_ms(path: str) -> int:
    """读音频时长（毫秒）。afinfo 优先，MP4 mvhd 兜底。"""
    try:
        out = subprocess.run(["afinfo", path], capture_output=True, text=True,
                             timeout=30).stdout
        m = re.search(r"estimated duration:\s*([\d.]+)\s*sec", out)
        if m:
            return int(round(float(m.group(1)) * 1000))
    except (OSError, subprocess.SubprocessError):
        pass
    return _mvhd_duration_ms(path)


def _mvhd_duration_ms(path: str) -> int:
    with open(path, "rb") as fh:
        data = fh.read(2 * 1024 * 1024)
    i = data.find(b"mvhd")
    if i < 0:
        raise ValueError(f"无法读取音频时长：{path}")
    ver = data[i + 4]
    if ver == 1:
        timescale, duration = struct.unpack(">IQ", data[i + 24:i + 36])
    else:
        timescale, duration = struct.unpack(">II", data[i + 16:i + 24])
    return int(round(duration / timescale * 1000))


def _next_rid(rels_xml: str) -> int:
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return max(ids, default=0) + 1


def _max_shape_id(slide_xml: str) -> int:
    ids = [int(n) for n in re.findall(r'<p:cNvPr id="(\d+)"', slide_xml)]
    if not ids:                                    # 退回 spTree 的 grpSpPr 之外
        ids = [int(n) for n in re.findall(r'<p:cNvPr[^>]*id="(\d+)"', slide_xml)]
    return max(ids, default=1)


def _audio_pic(spid: int, rid_audio: int, rid_media: int, rid_img: int) -> str:
    return (
        f'<p:pic><p:nvPicPr><p:cNvPr id="{spid}" name="Audio {spid - 1}"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr><a:audioFile r:link="rId{rid_audio}"/>'
        f'<p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}">'
        f'<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
        f' r:embed="rId{rid_media}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="rId{rid_img}"/><a:stretch><a:fillRect/>'
        f'</a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="{ICON_X}" y="{ICON_Y}"/>'
        f'<a:ext cx="{ICON_CX}" cy="{ICON_CY}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )


def _timing(spid: int) -> str:
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never"'
        ' nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        '<p:par><p:cTn id="3" fill="hold"><p:stCondLst>'
        '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/>'
        '</p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par>'
        '<p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0"'
        ' fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/>'
        '</p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)">'
        '<p:cBhvr><p:cTn id="6" dur="1" fill="hold"/><p:tgtEl>'
        f'<p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst>'
        '</p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
        '</p:par></p:childTnLst></p:cTn><p:prevCondLst>'
        '<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
        '</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl>'
        '<p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
        f'<p:audio isNarration="1"><p:cMediaNode vol="80000" showWhenStopped="0">'
        '<p:cTn id="7" fill="hold" display="0"><p:stCondLst>'
        '<p:cond delay="indefinite"/></p:stCondLst><p:endCondLst>'
        '<p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
        '</p:cond></p:endCondLst></p:cTn><p:tgtEl>'
        f'<p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:audio>'
        '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    )


def _add_default_ct(ct_xml: str, ext: str, mime: str) -> str:
    if re.search(rf'<Default Extension="{re.escape(ext)}"', ct_xml, re.I):
        return ct_xml
    node = f'<Default Extension="{ext}" ContentType="{mime}"/>'
    return ct_xml.replace("</Types>", node + "</Types>")


def _slide_part_names(zf: zipfile.ZipFile) -> list[str]:
    names = [n for n in zf.namelist()
             if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
    return sorted(names, key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))


def find_audio(audio_dir: str, slide_no: int) -> str | None:
    for ext in AUDIO_EXTS:
        for cand in (str(slide_no) + ext, str(slide_no) + ext.upper(),
                     f"{slide_no:02d}{ext}"):
            p = os.path.join(audio_dir, cand)
            if os.path.isfile(p):
                return p
    return None


def embed(deck: str, audio_dir: str, out: str) -> dict:
    zin = zipfile.ZipFile(deck)
    slides = _slide_part_names(zin)
    items = {}
    for part in slides:
        no = int(re.search(r"(\d+)", part.split("/")[-1]).group(1))
        path = find_audio(audio_dir, no)
        if path:
            items[part] = (no, path)

    if not items:
        raise SystemExit(f"在 {audio_dir} 里没找到 1.m4a / 2.m4a ... 这样的音频文件")

    poster_part = "ppt/media/narration_poster.png"
    payload = dict((n, zin.read(n)) for n in zin.namelist())
    zin.close()

    payload[poster_part] = _png_1x1_transparent()

    existing_media = [n for n in payload if n.startswith("ppt/media/")]
    ct_xml = payload["[Content_Types].xml"].decode("utf-8")
    report = []

    for part, (no, path) in sorted(items.items(), key=lambda kv: kv[1][0]):
        ext = os.path.splitext(path)[1].lower()
        if ext not in AUDIO_EXTS:
            raise SystemExit(f"不支持的音频格式：{path}")
        ms = audio_duration_ms(path)
        media_part = f"ppt/media/narration{no}{ext}"
        payload[media_part] = open(path, "rb").read()
        ct_xml = _add_default_ct(ct_xml, ext.lstrip("."), AUDIO_EXTS[ext])

        rels_part = f"ppt/slides/_rels/{os.path.basename(part)}.rels"
        rels = payload[rels_part].decode("utf-8")
        r = _next_rid(rels)
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="rId{r}" Type="{AUDIO_REL}" Target="../media/{os.path.basename(media_part)}"/>'
            f'<Relationship Id="rId{r + 1}" Type="{MEDIA_REL}" Target="../media/{os.path.basename(media_part)}"/>'
            f'<Relationship Id="rId{r + 2}" Type="{IMAGE_REL}" Target="../media/narration_poster.png"/>'
            "</Relationships>")
        payload[rels_part] = rels.encode("utf-8")

        xml = payload[part].decode("utf-8")
        spid = _max_shape_id(xml) + 1
        pic = _audio_pic(spid, r, r + 1, r + 2)
        xml = xml.replace("</p:spTree>", pic + "</p:spTree>", 1)
        # CT_Slide 顺序：cSld, clrMapOvr, transition, timing, extLst
        xml = xml.replace(
            "</p:sld>",
            f'<p:transition advTm="{ms}"/>' + _timing(spid) + "</p:sld>", 1)
        payload[part] = xml.encode("utf-8")
        report.append({"slide": no, "file": os.path.basename(path),
                       "seconds": round(ms / 1000, 1), "spid": spid})

    payload["[Content_Types].xml"] = ct_xml.encode("utf-8")

    app = payload["docProps/app.xml"].decode("utf-8")
    app = re.sub(r"<MMClips>\d+</MMClips>", f"<MMClips>{len(report)}</MMClips>", app)
    payload["docProps/app.xml"] = app.encode("utf-8")

    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in payload.items():
            zout.writestr(name, data)

    missing = [int(re.search(r"(\d+)", p.split("/")[-1]).group(1))
               for p in slides
               if p not in items]
    return {"out": out, "slides": len(slides), "embedded": report,
            "missing_audio_for_slides": sorted(missing),
            "total_seconds": round(sum(r["seconds"] for r in report), 1)}


def verify(deck: str) -> dict:
    z = zipfile.ZipFile(deck)
    rows = []
    for part in _slide_part_names(z):
        no = int(re.search(r"(\d+)", part.split("/")[-1]).group(1))
        xml = z.read(part).decode("utf-8")
        rels_part = f"ppt/slides/_rels/{os.path.basename(part)}.rels"
        rels = z.read(rels_part).decode("utf-8") if rels_part in z.namelist() else ""
        aud = re.search(r'<a:audioFile r:link="(rId\d+)"/>', xml)
        target = None
        if aud:
            m = re.search(rf'Id="{aud.group(1)}" Type="{re.escape(AUDIO_REL)}"'
                          r' Target="([^"]+)"', rels)
            target = m.group(1) if m else None
        adv = re.search(r'<p:transition[^>]*advTm="(\d+)"', xml)
        narr = 'isNarration="1"' in xml
        media_ok = bool(target) and (
            "ppt/" + target.replace("../", "")).replace("ppt/ppt/", "ppt/") in z.namelist()
        rows.append({"slide": no, "narration": narr,
                     "audio": target, "media_present": media_ok,
                     "advTm_ms": int(adv.group(1)) if adv else None})
    app = z.read("docProps/app.xml").decode("utf-8")
    mm = re.search(r"<MMClips>(\d+)</MMClips>", app)
    return {"deck": deck, "slides": len(rows), "MMClips": int(mm.group(1)) if mm else None,
            "tracks": rows}


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == "__main__":
    args = [a for a in sys.argv[1:]]
    if not args:
        raise SystemExit(__doc__)
    if args[0] == "--verify":
        import json
        print(json.dumps(verify(args[1]), ensure_ascii=False, indent=2))
    else:
        import json
        deck, adir = args[0], args[1]
        out = args[2] if len(args) > 2 else os.path.splitext(deck)[0] + "_已录音.pptx"
        print(json.dumps(embed(deck, adir, out), ensure_ascii=False, indent=2))
