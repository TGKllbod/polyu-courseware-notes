# 03 · 生成复习笔记 docx

## 铁律：一讲一个文件

```
✅ ABCD1234_复习笔记-第1讲.docx
✅ ABCD1234_复习笔记-第2讲.docx
❌ ABCD1234_复习笔记.docx          # 全学期合并成一个，期末翻不动
```

理由很实际：考前你只复习第 6–9 讲，合并版要滚两千行；而且每次改一讲都要重生成整本，风险大。

## 1. 总源长什么样

`复习笔记.md` 里，每讲用二级标题，讲号递增：

```markdown
## 9. Lecture 6 — 主题名（日期）

### 9.1 殖民地时期
…

## 10. Lecture 7 — 主题名（日期）

### 10.1 核心论断
…
```

`split_to_lectures.py` 就按 `## N. Lecture M — …` 切。**破折号必须是 `—`／`–`／`-`／`:`／`(` 之一**，否则切不出来。

## 2. 生成

```bash
cw notes ABCD1234_课程名
```

等价于：

```bash
python skills/polyu-courseware-notes/scripts/split_to_lectures.py \
  ~/PolyU/ABCD1234_课程名/复习笔记.md --prefix ABCD1234_
```

⚠️ **`--prefix` 必须传**。不传会生成裸的 `复习笔记-第1讲.docx`，和带课程代码的版本并存，早晚搞混。

## 3. 生成后必须检查四件事

| 检查 | 怎么查 |
| --- | --- |
| 只含该讲内容 | 打开第 3 讲，确认没有第 4 讲的小节 |
| 表格没丢 | 数一下表格数量，和 md 里对得上 |
| 无 md 语法残留 | 搜 `---`、`\| ---`、裸 `**` |
| 中文正常 | 标题、正文都不该出现方块或方框 |

```bash
cw check
```

`cw check` 会跑：
- `sync_check.py` —— 哪些 docx 比 md 旧（漏生成了）
- `acronym_check.py` —— 缩写首次出现有没有给全称

## 4. 缩写要给全称

第一次出现写 `CPF（Central Provident Fund，中央公积金）`，之后才能直接用 `CPF`。

```bash
python skills/polyu-courseware-notes/scripts/acronym_check.py 复习笔记.md
```

输出「合计: 0 处」才算过。

## 5. 表格转换的坑

`md_to_docx.py` 对**表格单元格内的嵌套强调**处理不干净，星号会漏进 Word：

```markdown
| 出处 |
| --- |
| **OECD, *Pensions at a Glance* 2023** |   ← 会漏成 **OECD, *Pensions...*
```

**表格里一律写纯文本**，强调放到正文段落。校验：

```python
from docx import Document
d = Document("out.docx")
bad = [(t, c.text) for t in d.tables for r in t.rows for c in r.cells if "*" in c.text]
print(bad)   # 必须是 []
```

## 6. 改动流程

```
改 复习笔记.md  →  cw notes  →  cw check  →  交付
```

**永远不要直接在 docx 上改**——下次生成就覆盖了。如果手上只有 docx 没有 md：

```bash
python skills/polyu-courseware-notes/scripts/docx_to_md.py 孤本.docx
```

## 下一步

→ [04-做PPT与讲稿](04-做PPT与讲稿.md)
