# PdfBilingualInspector

独立的 PDF 双语导出检查工具。

用途：
- 输入原始 `pdf`
- 读取对应的双语 `docx`
- 输出疑似未翻译片段报告
- 方便后续直接把报告作为 PDF 拆分/翻译修复的回归检查输入

这个工具不接入现有 WPF 主程序，不影响当前翻译功能。

## 用法

```powershell
dotnet restore tools\PdfBilingualInspector\PdfBilingualInspector.csproj --ignore-failed-sources /p:BaseIntermediateOutputPath='E:\translator\obj_verify_inspector\'
dotnet build tools\PdfBilingualInspector\PdfBilingualInspector.csproj --no-restore /p:BaseIntermediateOutputPath='E:\translator\obj_verify_inspector\' /p:OutDir='E:\translator\bin_verify_inspector\'
dotnet exec E:\translator\bin_verify_inspector\PdfBilingualInspector.dll --pdf "C:\path\paper.pdf"
```

也可以显式指定双语文档和报告前缀：

```powershell
dotnet exec E:\translator\bin_verify_inspector\PdfBilingualInspector.dll `
  --pdf "C:\path\paper.pdf" `
  --bilingual "C:\path\paper.bilingual.docx" `
  --report "E:\translator\tmp\paper-check"
```

## 输出

默认输出两份报告：
- `*.txt`：便于人工查看
- `*.json`：便于后续 Codex 程序化读取

默认报告路径为双语 `docx` 同目录下的：

```text
<原双语文件名>.inspection.txt
<原双语文件名>.inspection.json
```

## 返回码

- `0`：未发现明显未翻译片段
- `2`：发现疑似未翻译片段
- `1`：参数错误或运行失败

## 当前规则

当前会标记这些情况：
- 译文为空
- 译文与原文基本相同
- 译文中仍保留大段连续英文
- 译文几乎全是英文
- 译文看起来像直接保留了英文续句

说明：
- 这是启发式检查，不是语义级裁决
- 作者名、机构名、邮箱、编号等场景可能仍有少量误报
- 设计目标是优先帮助快速发现 PDF 拆分失败导致的明显漏译片段
