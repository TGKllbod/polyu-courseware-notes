---
name: ppt-authoring
description: 做或改 .pptx（尤其是按模板做）时，**必须用本机已安装的 LibreOffice 把 PPT 打开/渲染出来实际看一遍**再交付。触发情景：写或改 PPT、按模板排版、用户反馈「版面乱／文字看不清／不像模板」。
---

# PPT：动手改完，先用 LibreOffice 看一遍

本机装了 LibreOffice（`/Applications/LibreOffice.app/Contents/MacOS/soffice`）。**PPT 不能只靠读代码或想象版面就交付，必须渲染出来看。**

```bash
HOME="$PWD/.render" /Applications/LibreOffice.app/Contents/MacOS/soffice \
  -env:UserInstallation=file://<工作区>/.render/lo_profile \
  --headless --norestore --convert-to pdf --outdir <输出目录> <文件.pptx>
```

再把 PDF 逐页渲成图片，逐页看：**文字有没有溢出／压住别的东西、有没有重叠、有没有跨出模板色块、有没有模板残留文案、图片位是不是空着、各页疏密是否一致、有没有错别字。**

- `HOME` 与 `-env:UserInstallation` 都要指到工作区内，否则 soffice 会往用户目录写配置。
- LibreOffice **不还原** PPT 内嵌字体，字体看着和 PowerPoint 不同属正常；要比字体就用**模板原件**在同一渲染器里做基准。
- 用户已自行修改过的页面**不要覆盖**（改前备份、改后用形状指纹断言该页未变）。
