# 课程资料整理工作流（Courseware Notes Workflow）

一套用 AI 助手整理**整学期课件**、按讲生成**复习笔记**、写**演讲稿**与**小组作业**的可复制工作流。
它不是打包好的软件，而是**一套规则（skill）＋一组脚本**：你把它们放进自己的 AI 助手，就能用同样的方式管理一学期的材料。

> 最初为香港理工大学（PolyU）课程整理而写，方法本身与学校无关。

---

## 它解决什么问题

| 痛点 | 这套工作流的做法 |
| --- | --- |
| 一学期几十份课件散在下载文件夹 | 每门课一个文件夹，课件按原名归档，课级信息进 `README.md` |
| 期末复习要从头翻 PPT | 一份 `复习笔记.md` 作总源，**按讲**导出 `科目_复习笔记-第N讲.docx`（严禁合并成一个大文档） |
| 小组作业分工靠聊天记录 | 分工表、接口句、数据口径表落成文件，谁讲哪节一目了然 |
| 数字前后矛盾被打回 | 「共用数字口径表」把每个数字的口径、年份、出处、**常见错版**固定下来 |
| 演讲稿写成书面语、念起来别扭 | 「能念」硬指标：平均句长、长串上限、书面腔词表清零 |

## 核心原则（这套工作流的灵魂）

1. **一个可维护总源，多个成品**：`.md` 是总源，`.docx` 是交付物；改内容先改 md，再重新生成 docx。
2. **表格必须按单元格读**。用文本流顺序推断行列归属是这套方法里代价最高的错误——它会把「星期四的课」记成星期三，并让后面所有结论都错。
3. **不编造**：信息缺失就写「未提供」，多来源冲突就并列两种读法并标注待确认。
4. **推断要分层标注**：哪些是课件原话、哪些是 AI 的推断、哪些需要外部资料才能成立，逐条标明。
5. **交付前机械校验**：跑一遍检查脚本，而不是靠记忆确认「应该没错」。

## 快速开始

```bash
# 1) 依赖
pip install python-docx python-pptx pdfplumber pymupdf Pillow numpy

# 2) 指到你的课程仓库
export POLYU_HOME="$HOME/PolyU"
mkdir -p "$POLYU_HOME"

# 3) 把 skill 放进 AI 助手的 skills 目录
cp -R skills/* ~/.claude/skills/

# 4) 建一门课，开始整理
mkdir -p "$POLYU_HOME/ABCD1234_课程名"
# 把课件放进去，然后让助手按 skill 处理
```

**可选依赖**：LibreOffice（把 PPTX 渲染成 PDF 以便实际查看效果）、Tesseract 或 macOS Vision（图片型课件 OCR）。

### 用命令行（推荐）

```bash
pip install -e .          # 安装后获得 cw 命令
cw doctor                 # 先自检环境
cw init --root ~/PolyU --courses "ABCD1234:课程名"
cw ingest ~/Downloads/Lecture3.pptx
cw notes ABCD1234_课程名
cw check
```

不想装也可以直接跑：`python cw.py doctor`。

## 文档

按顺序读，每篇都能单独用：

| # | 文档 | 什么时候看 |
| --- | --- | --- |
| 01 | [第一次使用](docs/01-第一次使用.md) | 刚 clone 下来 |
| 02 | [整理一份课件](docs/02-整理一份课件.md) | 下载了新的 pptx / pdf |
| 03 | [生成复习笔记 docx](docs/03-生成复习笔记docx.md) | 要按讲导出 Word |
| 04 | [做 PPT 与讲稿](docs/04-做PPT与讲稿.md) | 要做展示、上台念稿 |
| 05 | [小组作业](docs/05-小组作业.md) | 要分工、要对接口、要保证数字一致 |
| 06 | [Windows 环境](docs/06-Windows环境.md) | 在 Windows 上装依赖、OCR、LibreOffice |

## CLI：`cw`

| 命令 | 作用 |
| --- | --- |
| `cw init --root ~/PolyU --courses "ABCD1234:课程名"` | 建工作区骨架：课程文件夹 + `README.md` 模板 |
| `cw doctor` | 环境自检：依赖、LibreOffice、OCR 后端、中文字体、git |
| `cw ingest <文件...> [--course 文件夹] [--extract] [--move] [--dry-run]` | 按文件名里的课程代码（`ABCD1234`、`MNOP3456` 都认）归档课件；`--extract` 顺带做四层提取 |
| `cw notes <课程文件夹> [--lecture N]` | 从 `复习笔记.md` 按讲拆分生成 docx |
| `cw check [--fix]` | 同步检查（md 比 docx 新）＋ 缩写检查 |
| `cw render <deck.pptx>` | LibreOffice 转 PDF，强制隔离用户配置目录 |
| `cw ocr <图片>` | macOS 走 Vision（自动编译）；其他平台走 Tesseract |

`--root` 写在子命令前后都可以：`cw --root X init` 与 `cw init --root X` 等价。
工作区根目录的解析顺序：`--root` → `$POLYU_HOME` → `~/PolyU`。

## Windows 支持

本工具链原本在 macOS 上开发，CLI 已针对 Windows 处理了以下差异：

| 问题 | 处理方式 |
| --- | --- |
| 控制台默认 `cp936`，中文会崩 | 启动时 `sys.stdout.reconfigure(encoding="utf-8")` |
| LibreOffice 路径不同 | `cw` 自动搜索 `C:\Program Files\LibreOffice\program\soffice.exe`（含 x86 目录）与 `PATH` |
| `-env:UserInstallation=file://` 的盘符写法 | 用 `Path.as_uri()` 生成 `file:///C:/...`，macOS/Linux 同样正确 |
| 隔离用的 `HOME` 变量在 Windows 无效 | 同时设置 `HOME` 与 `USERPROFILE`、`TMPDIR` |
| 图片 OCR 无 Vision | 回退到 Tesseract：`winget install UB-Mannheim.TesseractOCR` |
| 中文字体 | Windows 自带 Microsoft YaHei；macOS 回退 PingFang SC |
| 终端 emoji 显示异常 | CLI 输出统一用 ASCII 标记 `[OK] [!] [x]` |

安装 LibreOffice：`winget install TheDocumentFoundation.LibreOffice`

## 仓库结构

```
cw.py                          命令行入口（跨平台）
pyproject.toml                 pip install -e . 后得到 cw 命令
skills/
  polyu-courseware-notes/
    SKILL.md                    完整工作流规则（三层提取、表格铁律、docx 生成、时间线判定…）
    references/邮件语气样本.md    让 AI 学会你的邮件语气（方法＋规则，不含真实邮件）
    scripts/
      extract_pptx.py           PPTX 四层提取：坐标排序文本 + 逐格表格 + 图片 + 备注 + 图表数据
      md_to_docx.py             md → Word（标题/表格/列表/粗体/引用）
      split_to_lectures.py      按讲拆分，一讲一个 docx
      table_qa.py               表格结构化读取与合计验算
      acronym_check.py          缩写首现是否给全称
      sync_check.py             md 比 docx 新 → 报告过期项
      docx_to_md.py             docx 反导出 md（救回「只有 docx 没有源文件」的孤本）
      embed_narration.py        逐页录音旁白封装进 pptx
      pdf_render.swift          PDF 渲染成图（macOS）
      vision_ocr.swift          离线 OCR（macOS Vision）
  ppt-authoring/
    SKILL.md                    PPT 改完必须渲染出来看一遍：命令、隔离方式、常见坑
```

## 脚本速查

| 脚本 | 用途 |
| --- | --- |
| `extract_pptx.py <file.pptx> [--from N --to M]` | 一次拿到文本、表格（逐格）、图片、备注、图表数据 |
| `md_to_docx.py <in.md> [out.docx]` | 单文件转换；**不要无参数运行** |
| `split_to_lectures.py <课程>/复习笔记.md --prefix <代码>_` | 按讲拆分成多个 docx |
| `table_qa.py pdf\|pptx\|docx <file>` | 看表头坐标、逐格内容、合计是否成立 |
| `acronym_check.py <文件或目录>` | 缩写首现缺全称的清单 |
| `sync_check.py [--fix]` | 哪些 docx 已过期 |
| `docx_to_md.py <file.docx>` | 反导出 md |
| `embed_narration.py <deck.pptx> <audio_dir>` | 把逐页录音嵌进 pptx |

## 适合谁

- 一学期十几门课、材料堆在手里却从没整理过的学生；
- 想让 AI 持续维护同一套笔记，而不是每次从零问一遍的人；
- 愿意在交付前跑一遍检查脚本、接受「先校验再声称完成」的人。

**不适合**：只想要一个「一键生成论文」工具的人——这套方法不生成内容，它**组织证据、暴露矛盾、逼你标注不确定性**。

## 路线图

- [x] 规则与脚本（当前版本）
- [x] CLI：`cw init` / `doctor` / `ingest` / `notes` / `check` / `render` / `ocr`（含 Windows 适配）
- [ ] 本地 Web 界面：拖入课件 → 自动提取 → 预览笔记 → 导出 docx
- [ ] 模板化：不同学校/学科的 `复习笔记` 模板

## 隐私与版权（重要）

- **不要提交课程材料**：讲义、试卷、课程大纲的著作权属于教师/学校，仓库只放**方法**与**脚本**。
- **不要提交个人数据**：学号、成绩、同学姓名、教师私人邮箱、邮件原文。
- 本仓库**不含任何真实课件、成绩或个人身份信息**——没有姓名、学号、邮箱、教师或同学姓名，也没有带日期的个人活动记录。
- 文中出现的课程代码（QRST7890 等）**只是案例举例**，用于说明不同课程的结构差异。
- `references/` 只保留方法说明，不含任何真实邮件原文。
- 你自己的工作区应当**保持私有**。

## License

MIT
