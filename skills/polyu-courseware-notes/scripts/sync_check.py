# -*- coding: utf-8 -*-
"""检查「md 源 → docx 成品」是否同步（防交付物悄悄过期）。

背景（教训）：
    工作区里的《<课程代码>_课程总结.docx》《课程总览.docx》原先**只有 docx、没有 md 源**，
    是「孤本」。结果源信息更新了好几轮（10/22 补课、11/12 报告、课室变更……），
    这些 docx 却一直停在 9 月中旬——用户翻《课程总结》时发现对不上，才知道漏了。
    现在它们都有了 .md 源，本脚本用**修改时间**做机械校验：**md 比 docx 新 ⇒ docx 已过期**。

用法:
    python3 sync_check.py              # 检查整个工作区
    python3 sync_check.py --fix        # 顺手把过期项重新生成（调用 md_to_docx.py）

判定规则:
    对每个存在同名 .docx 的 .md，要求 docx 的 mtime >= md 的 mtime。
    （md_to_docx.py 生成 docx 后其 mtime 自然更新，所以正常同步时不会报。）
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def _find_ws() -> str:
    """确定工作区根目录。优先级：POLYU_HOME > 从 cwd 向上找 > 脚本位置回退。

    不能只靠脚本位置：skill 被装到 ~/.claude/skills/ 时，向上 4 层会变成家目录。
    """
    env = os.environ.get("POLYU_HOME")
    if env:
        return os.path.abspath(os.path.expanduser(env))
    marks = ["课程总览.md", "DDL汇总.md", "复习笔记.md"]
    here = os.getcwd()
    for d in [here] + [os.path.dirname(p) for p in _ancestors(here)]:
        if any(os.path.exists(os.path.join(d, m)) for m in marks):
            return d
        # 也认「课程文件夹」形态：根下某个子目录里有 复习笔记.md
        try:
            for sub in os.listdir(d):
                if os.path.exists(os.path.join(d, sub, "复习笔记.md")):
                    return d
        except OSError:
            pass
    # 回退：假定脚本在 <ws>/.claude/skills/<skill>/scripts/
    return os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))


def _ancestors(path):
    p = os.path.abspath(path)
    out = []
    while True:
        parent = os.path.dirname(p)
        if parent == p:
            return out
        out.append(parent)
        p = parent


WS = _find_ws()

# 需要成对维护的（md → docx）
PAIRS = [
    ("课程总览.md", "课程总览.docx"),
    ("全部课程排表.md", "全部课程排表.docx"),
    ("DDL汇总.md", "DDL汇总.docx"),
    ("出勤规定汇总.md", "出勤规定汇总.docx"),
    ("AI使用红线速查.md", "AI使用红线速查.docx"),
]
COURSES = [
    "QRST7890_Social_Policy_Analysis_and_Social_Accountability",
    "ABCD1234_Policy_Research_Method",
    "EFGH5678_Comparative_Social_Policy_and_Social_Development",
    "IJKL9012_Theory_and_Practice_of_Social_Development",
    "MNOP3456_Academic_Integrity_and_Ethics",
]
for c in COURSES:
    code = c.split("_")[0]
    PAIRS.append((f"{c}/{c.split('_')[0]}_课程总结.md", f"{c}/{code}_课程总结.docx"))

SKIP_DIRS = {".render", ".pylibs", ".claude", "__pycache__"}


def newest_docx_for(md_path):
    """复习笔记.md 对应的是按讲拆分的多个 docx，单独处理。"""
    return None


def check(fix=False):
    stale, missing, ok = [], [], []
    for md_rel, dx_rel in PAIRS:
        md = os.path.join(WS, md_rel)
        dx = os.path.join(WS, dx_rel)
        if not os.path.exists(md):
            continue
        if not os.path.exists(dx):
            missing.append((md_rel, dx_rel))
            continue
        m_md, m_dx = os.path.getmtime(md), os.path.getmtime(dx)
        if m_dx + 1 < m_md:          # 容 1 秒
            stale.append((md_rel, dx_rel, m_md - m_dx))
        else:
            ok.append(dx_rel)

    # 复习笔记：md → 各讲 docx（只要有一个讲次 docx 比 md 旧就算过期）
    for c in COURSES:
        code = c.split("_")[0]
        md = os.path.join(WS, c, "复习笔记.md")
        d = os.path.join(WS, c)
        if not (os.path.exists(md) and os.path.isdir(d)):
            continue
        m_md = os.path.getmtime(md)
        lect = [f for f in os.listdir(d)
                if f.startswith(f"{code}_复习笔记-第") and f.endswith(".docx")]
        if not lect:
            missing.append((f"{c}/复习笔记.md", f"{code}_复习笔记-第N讲.docx"))
            continue
        aged = [(f, m_md - os.path.getmtime(os.path.join(d, f)))
                for f in lect if os.path.getmtime(os.path.join(d, f)) + 1 < m_md]
        if aged:
            worst = max(x[1] for x in aged)      # 最旧的那一个 docx
            stale.append((f"{c}/复习笔记.md",
                          "、".join(sorted(f for f, _ in aged)), worst))
        else:
            ok.append(f"{c}/ 各讲 docx（{len(lect)} 个）")

    print(f"工作区: {WS}\n")
    if ok:
        print(f"✓ 已同步（{len(ok)}）")
        for x in ok:
            print(f"    {x}")
    if stale:
        print(f"\n✗ **已过期**（md 比 docx 新，docx 未重新生成）（{len(stale)}）")
        for md_rel, dx_rel, delta in stale:
            print(f"    {dx_rel}")
            print(f"        ← 源 {md_rel} 新了约 {delta/3600:.1f} 小时")
    if missing:
        print(f"\n⚠ 缺 docx（{len(missing)}）")
        for md_rel, dx_rel in missing:
            print(f"    {md_rel} → 应有 {dx_rel}")
    if not stale and not missing:
        print("\n结论: 全部同步 ✓")
    else:
        print(f"\n结论: {len(stale)} 项过期、{len(missing)} 项缺件"
              f"—— 请对每个 .md 运行：python3 scripts/md_to_docx.py <该 md>")

    if fix and stale:
        print("\n--fix：重新生成 --------------------------------------------------")
        script = os.path.join(HERE, "md_to_docx.py")
        for md_rel, dx_rel, _ in stale:
            if md_rel.endswith("复习笔记.md"):
                continue   # 复习笔记须用 split_to_lectures.py，跳过
            r = subprocess.run([sys.executable, script, os.path.join(WS, md_rel)],
                               capture_output=True, text=True)
            print(f"  {r.stdout.strip() or r.stderr.strip()}")
    return 1 if (stale or missing) else 0


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == "__main__":
    sys.exit(check(fix="--fix" in sys.argv))
