#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
缩写检查器（配合 SKILL.md §二第 8 条）。
规则：英文缩写**首次出现**必须有全称或解释，例如：
    Import-substitution industrialization (ISI)   /   ISI（进口替代工业化）
    Youth Employment Initiative (YEI)             /   YEI（青年就业倡议）
检查逻辑：扫出所有 2-6 个大写字母的缩写，逐个看**首次出现**处前后是否有
全称或中文解释（ASCII/中文括号内含 ≥4 字符，或紧跟/前置全称），无则报告。

用法:
    python3 acronym_check.py <文件或目录> [<更多路径>...]
    python3 acronym_check.py ../../APSS5521_*/复习笔记.md
"""
import os
import re
import sys

# 极常见、通常无需解释（同一文档若影响理解仍建议加）
SAFE = {
    'HK', 'US', 'UK', 'UN', 'EU', 'AI', 'PDF', 'PPT', 'PPTX', 'DDL', 'GDP', 'GNP', 'HDI',
    'NGO', 'APA', 'OECD', 'ILO', 'WHO', 'WTO', 'IMF', 'TV', 'PC', 'URL', 'OK', 'ID', 'AR',
    'APSS', 'HTI', 'MASPSD', 'H0', 'H1', 'H2', 'H3', 'S1', 'S2', 'PhD', 'Dr', 'Prof',
    'W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8', 'W9', 'W10', 'W11', 'W12', 'W13',
    # 提取脚本的输出标记（不是缩写，出现在关于 extract_pptx.py 的说明文字里）
    'TABLE', 'NOTES', 'SLIDE',
    # 化学式（CO₂ 的下标字符不在 [A-Z] 内，会被切成 CO）
    'CO', 'CO2', 'PM',
    # 立法会文件编号前缀（如 CB(2)233/17-18(06)），非缩写
    'CB',
    # HTTP 方法名（如"读取网页的能力（GET）"），非缩写
    'GET',
}
# 纯强调性大写词，不是缩写
NOT_ACRONYM = {
    'NOT', 'NO', 'MAY', 'AND', 'OR', 'ONE', 'TWO', 'TOO', 'STEP', 'PART', 'AND', 'ALL',
    'THE', 'II', 'III', 'IV', 'V', 'VI', 'FIRST', 'SECOND', 'THIRD', 'MUST', 'YES', 'ONLY',
    'BINARY', 'REALITY', 'POLICY', 'DATA', 'METHODS', 'MODEL', 'CASES', 'KEY', 'OCR_HINT',
    # 引文中的强调性大写（课件的斜体/加粗效果），如 "those who can do something about it WILL do so"
    'WILL', 'MUST', 'DOES', 'LOOKS', 'CAN', 'NOT', 'NOW', 'HERE', 'ONLY',
    # 引用官方原文时的强调性大写，如 "must be YOUR OWN WORK"
    'YOUR', 'OWN', 'WORK',
    # 声明模板里的占位符（如 "[NAME OF TOOL] … in order to [REASON]"），非缩写
    'REASON', 'NAME', 'TOOL', 'SERVICE', 'TITLE',

    "NEXT",          # 讲稿的翻页提示词，非缩写
    "FOR", "WOMEN",  # 全大写英文标题里的普通词（APSS5521 L7 p17「DEVELOPMENT WORKING FOR WOMEN」），非缩写
    'NOT', 'NO', 'MAY', 'AND', 'OR', 'ONE', 'TWO', 'TOO', 'STEP', 'PART', 'AND', 'ALL',
    'THE', 'II', 'III', 'IV', 'V', 'VI', 'FIRST', 'SECOND', 'THIRD', 'MUST', 'YES', 'ONLY',
    'BINARY', 'REALITY', 'POLICY', 'DATA', 'METHODS', 'MODEL', 'CASES', 'KEY', 'OCR_HINT',
    # 引文中的强调性大写（课件的斜体/加粗效果），如 "those who can do something about it WILL do so"
    'WILL', 'MUST', 'DOES', 'LOOKS', 'CAN', 'NOT', 'NOW', 'HERE', 'ONLY',
    # 引用官方原文时的强调性大写，如 "must be YOUR OWN WORK"
    'YOUR', 'OWN', 'WORK',
    # 声明模板里的占位符（如 "[NAME OF TOOL] … in order to [REASON]"），非缩写
    'REASON', 'NAME', 'TOOL', 'SERVICE', 'TITLE',

    "NEXT",          # 讲稿的翻页提示词，非缩写
}

ACRO = re.compile(r'(?<![A-Za-z0-9&])([A-Z]{2,6}s?)(?![A-Za-z])')


def is_glossed(text, start, acro):
    """首次出现处附近是否有全称或解释。"""
    window = text[max(0, start - 80):start + len(acro) + 100]
    # 括号内含中文或英文全称（≥4 字符）
    if re.search(r'[（(][^）)\n]{4,}[）)]', window):
        return True
    # 英文全称紧跟括号缩写：Full Form (ACRO)
    if re.search(r'[A-Za-z][A-Za-z\- ]{5,}[（(]' + re.escape(acro.rstrip('s')) + r's?[）)]', window):
        return True
    # 常见连接词提示
    if re.search(r'(全称|即|简称|也就是|namely|stands for|also known as)', window):
        return True
    return False


def check_file(path):
    text = open(path, encoding='utf-8').read()
    seen = {}
    for m in ACRO.finditer(text):
        acro = m.group(1)
        base = acro.rstrip('s')
        if base in SAFE or acro in SAFE or base in NOT_ACRONYM:
            continue
        tail = text[m.end():m.end() + 12]
        head = text[max(0, m.start() - 3):m.start()]
        # 误报过滤：教室/房间代码（ACRO+3位数字）、文件名、姓名（ACRO + 空白 + 首字母大写的名字）
        if re.match(r'\d{3}', tail):
            continue
        if acro in ('README', 'MD', 'DOCX', 'XLSX'):
            continue
        # 误报过滤：是文件名的一部分（ACRO 紧跟扩展名，如 SKILL.md、NOTES.py、TABLE.docx）
        if re.match(r'\.(md|py|docx|doc|pptx|ppt|xlsx|pdf|txt|json|swift|html|csv)\b', tail, re.I):
            continue
        if re.match(r'\s+[A-Z][a-z]+', tail) and re.search(r'[\u4e00-\u9fff]', text[m.end():m.end() + 20]):
            continue
        if head.strip().endswith('·'):
            continue
        if acro in seen:
            continue
        if is_glossed(text, m.start(), acro):
            seen[acro] = None            # 首次出现已解释
        else:
            ctx = text[max(0, m.start() - 70):m.start() + 70].replace('\n', ' ')
            seen[acro] = ctx.strip()
    return {a: c for a, c in seen.items() if c}


def main():
    paths = sys.argv[1:] or ['.']
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, names in os.walk(p):
                files += [os.path.join(root, n) for n in names if n.endswith(('.md',))]
        else:
            files.append(p)
    total = 0
    for f in sorted(set(files)):
        bad = check_file(f)
        if not bad:
            continue
        total += len(bad)
        print(f'== {f} | {len(bad)} 个缩写首次出现处缺全称/解释')
        for a, c in sorted(bad.items()):
            print(f'   {a:10s} … {c[:110]}')
    print(f'\n合计: {total} 处（SAFE 列表内的常见缩写已跳过）')
    return 1 if total else 0


try:  # Windows 控制台默认 cp1252/cp936，统一转 UTF-8，否则打印中文会 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if __name__ == '__main__':
    sys.exit(main())
