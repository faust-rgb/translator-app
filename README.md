# TranslatorApp

本项目是一个本地运行的 AI 文档翻译器，面向 Windows 桌面环境。

## 已实现能力

- 可手动配置大模型访问地址、模型名、API Key、自定义请求头。
- 当前支持 `AnthropicCompatible`，并预留 `OpenAiCompatible` 协议。
- 支持批量加入 `DOCX / XLSX / PPTX / PDF`。
- Office 文档通过 OpenXML 在本地改写，只向模型发送文本内容。
- PDF 按页读取，并以文本块级重绘方式替换译文，较之前的逐行覆盖更稳定。
- 支持输出目录、输出字体、字号、日志、总进度、暂停、恢复、停止。
- 支持流式翻译预览。
- 支持术语表文件，按 `原文=译文` 或 `原文<TAB>译文` 编写。
- 支持术语表热加载，修改术语表后新任务会自动使用最新内容。
- 支持失败自动重试。
- 支持为每个源文档额外导出一份双语对照 `DOCX`。
- 支持任务历史持久化。
- 对扫描版 PDF 增加 OCR 回退模式，需要本地准备 `tessdata`。
- OCR 分段已增强为“词 -> 行 -> 栏 -> 段落块”聚合，对复杂版式扫描 PDF 更稳。
- 支持崩溃恢复与断点续跑，程序重启后会自动恢复未完成任务。
- 可通过 `publish.ps1` 打包为单文件 EXE。
- 可通过 `installer.iss` 与 `build-installer.ps1` 生成安装包。

## 运行

```powershell
dotnet run --project .\TranslatorApp\TranslatorApp.csproj
```

## 打包

```powershell
.\publish.ps1
```

打包结果默认输出到 `.\publish`。

## OCR 准备

将 `eng.traineddata`、`chi_sim.traineddata` 放到 `.\tessdata`，或者在界面中指定 `tessdata` 目录。

## 术语表

不需要下载专门的术语表，直接使用项目里的示例文件即可：

- [glossary-template.txt](/e:/translator/glossary-template.txt)

你可以在这个文件基础上继续追加自己的术语。

## 安装包

安装 Inno Setup 后执行：

```powershell
.\build-installer.ps1
```
