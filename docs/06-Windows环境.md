# 06 · Windows 环境

## 1. 装 Python 与依赖

从 [python.org](https://www.python.org/downloads/windows/) 装 3.9+（勾选 **Add python.exe to PATH**）。

```powershell
git clone https://github.com/TGKllbod/polyu-courseware-notes.git
cd polyu-courseware-notes
pip install -e .
cw doctor
```

## 2. LibreOffice（`cw render` 需要）

```powershell
winget install TheDocumentFoundation.LibreOffice
```

或者从 [libreoffice.org](https://www.libreoffice.org/download/) 下载。装完**重开终端**再跑 `cw doctor`。

`cw` 会自动找到：

```
C:\Program Files\LibreOffice\program\soffice.exe
C:\Program Files (x86)\LibreOffice\program\soffice.exe
```

## 3. OCR：图片型课件

macOS 有系统自带的 Vision，Windows 没有。装 Tesseract：

```powershell
winget install UB-Mannheim.TesseractOCR
```

装完重开终端，`cw doctor` 应显示 `OCR：tesseract：...`。

之后：

```powershell
cw ocr slide5.png
```

## 4. 中文乱码：这是最容易踩的坑

Windows 控制台默认编码是 `cp1252` 或 `cp936`，脚本一打印中文就：

```
UnicodeEncodeError: 'charmap' codec can't encode characters in position 78-81
```

本仓库已处理两层：

- `cw` 调用子脚本时注入 `PYTHONIOENCODING=utf-8` 与 `PYTHONUTF8=1`
- 8 个脚本各自在启动时 `sys.stdout.reconfigure(encoding="utf-8")`

**如果你自己写了新脚本**，记得加同样的守卫：

```python
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
```

另外，Windows 终端对 emoji 支持不稳，CLI 输出统一用 `[OK] [!] [x]` 这类 ASCII 标记。

## 5. 路径与盘符

`cw render` 需要给 LibreOffice 传 `file:///` 形式的 URI。手写容易错，尤其盘符：

| 系统 | 正确写法 |
| --- | --- |
| Windows | `file:///C:/Users/you/PolyU/.render/lo_profile` |
| macOS | `file:///Users/you/PolyU/.render/lo_profile` |

`cw` 用 `Path.as_uri()` 自动处理。自己写脚本时不要手拼字符串，尤其路径里有空格或中文。

## 6. 环境变量

```powershell
setx POLYU_HOME "%USERPROFILE%\PolyU"
```

**重开终端**生效。之后 `cw init`、`cw notes` 都不用带 `--root`。

## 7. 已由 CI 覆盖的组合

| 系统 | Python 3.9 | Python 3.12 |
| --- | --- | --- |
| Windows | ✅ | ✅ |
| macOS | ✅ | ✅ |
| Ubuntu | ✅ | ✅ |

CI 会跑：`pip install -e .` → `cw --help` → `cw doctor` → `cw init` → `cw notes` → 检查 docx 生成。
所以上面这些适配不是「理论上应该行」，是每次提交都验过的。
