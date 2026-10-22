# 04 · 做 PPT 与讲稿

## 一、PPT：改完必须渲染出来看

坐标算对了 ≠ 好看。文字压到花纹上、两行叠在一起，坐标检查看不出来。

```bash
cw render deck.pptx
```

生成 `deck.pdf`，然后用阅读器实际翻一遍。

### 为什么不能直接 `soffice` 裸跑

无图形界面环境下会报 `DeploymentException`，因为 LibreOffice 想写默认用户目录却写不进去。必须隔离：

```bash
soffice -env:UserInstallation=file:///path/to/.render/lo_profile \
        --headless --norestore --convert-to pdf --outdir out deck.pptx
```

`cw render` 已经做了这件事，并处理了 Windows 的 `file:///C:/...` 与 `HOME`/`USERPROFILE`。

### 三个高频坑

| 现象 | 原因 | 解法 |
| --- | --- | --- |
| 渲染出来字体全变了 | PowerPoint 内嵌字体会被混淆，LibreOffice 解不开 | 正常现象。别因此改字体，改完在 PowerPoint 里看 |
| 用形状名称批量删元素，装饰全没了 | `has_text_frame` 会把 freeform 也算进去 | 只认 `shape_type == TEXT_BOX` |
| 改完再跑一次，元素越积越多 | 脚本不幂等 | 重画前先 `clear_own()` 删掉自己上次加的 |

**验证要验证到行**：把 PDF 里每个文本块的 bbox 拿出来，两两比 `dx` 与 `dy`。只比 `dy` 会把并排的两栏误判成重叠。

## 二、讲稿：写给耳朵，不是写给眼睛

### 硬指标

- 平均句长 ≤ 11 词/句，**不得有超过 22 词的句子**
- 不得有超过 13 词的无标点长串
- 书面腔词表清零：`such as`（口语用 like）、`shall`、`thereby`、`in terms of`、`with regard to`、`utilise`

### 时长换算

| 场景 | 语速 | 2 分钟 ≈ |
| --- | --- | --- |
| 学术报告，念稿 | 130–150 wpm | 280–310 词 |
| 学术报告，舒适 | 100–130 wpm | 200–260 词 |

**先写再念一遍计时**，别按字数估算就上台。

### 七类「念出来很怪异」的病因

1. **逗号插入语当绊脚石** —— `personal data, once collected, cannot be taken back` → `once personal data is collected, it cannot be taken back`
2. **超长名词串** —— 机构全名塞进列表首项，听众断线。降级成 `China's cyberspace regulator`
3. **列表项长短悬殊** —— 四项里第一项 11 词、其余 3 词，节奏崩
4. **被动语态堆叠** —— 能改主动就改
5. **抽象名词当动词用** —— `records of what is kept must be retained` → `companies should keep a record`
6. **连接词反复** —— 连续几句都用 `And` 开头会腻（但短句间用 `And` 起头反而自然，留少量）
7. **同段重复同一个词** —— 论点句结尾换个说法

### 分页与翻页

讲稿要标出翻页点，并**一句一页地对齐幻灯片**：

```
SLIDE 3 — Research questions          0:40

So the question we ask is this: …

Next
```

`Next` 独占一行，**念的时候不用念出来**——它是给你自己的翻页提示。

### 最后：回扣开头

开头立的对立、提的问题，结尾要收回来。一句 `which answers the question we opened with` 就能让听众觉得你讲的是一个论证，而不是五段资料。

## 三、如果你有书面报告

**报告原句优先，绝不自行改写。** 讲稿照抄报告，只允许四类朗读必需的改动：

1. 分号 → 句号
2. 去掉括号与引号
3. 标题式冒号 → 句号
4. 加最小过渡句

然后做保真校验：报告段落逐句在讲稿中做子串匹配，多重集词差必须为空。

## 下一步

→ [05-小组作业](05-小组作业.md)
