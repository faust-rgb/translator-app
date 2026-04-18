# TranslatorApp

本项目是一个本地运行的 AI 文档翻译器，面向 Windows 桌面环境。

## 已实现能力

- 可手动配置大模型访问地址、模型名、API Key、自定义请求头。
- 当前支持 `AnthropicCompatible`，并预留 `OpenAiCompatible` 协议。
- 支持批量加入 `DOCX / XLSX / PPTX / PDF / EPUB / MOBI / AZW3`。
- Office 文档通过 OpenXML 在本地改写，只向模型发送文本内容。
- PDF 按页读取，并以文本块级重绘方式替换译文，较之前的逐行覆盖更稳定。
- PDF 翻译已加入段落块重组、跨块断词修复、页边噪声过滤、公式保留、块类型识别、OCR 过滤和中文字符级换行，适合论文类 PDF。
- 公式检测已区分“纯公式块”和“含公式正文”：纯公式保持原样，含公式正文会继续翻译并提示保留公式/变量/编号。
- 对 PDF 英文残片、跨页续句和断词续句增加了二次重试与上下文提示，降低正文漏译概率。
- 支持输出目录、输出字体、字号、日志、总进度、暂停、恢复、停止。
- 支持流式翻译预览，预览区会自动换行。
- 支持术语表文件，按 `原文=译文` 或 `原文<TAB>译文` 编写。
- 支持术语表热加载，修改术语表后新任务会自动使用最新内容。
- 支持失败自动重试。
- 支持为每个源文档额外导出一份双语对照 `DOCX`。
- 支持任务历史持久化。
- 对扫描版 PDF 增加 OCR 回退模式，需要本地准备 `tessdata`。
- OCR 分段已增强为“词 -> 行 -> 栏 -> 段落块”聚合，对复杂版式扫描 PDF 更稳。
- 支持崩溃恢复与断点续跑，程序重启后会自动恢复未完成任务。
- 支持独立配置文档级并发、块级并发和全局远端请求并发；默认采用安全模式 `1/1/1`，避免远端 AI 并发限流导致任务失败。
- Word 翻译会按连续相同格式分组回填译文，尽量保留 Run 级格式边界。
- EPUB 会在本地按 XHTML / TOC 结构翻译，尽量保留章节、图片、内联样式与目录结构。
- 电子书现可原生输出为 `EPUB` 或 `DOCX`。
- 电子书导出 `DOCX` 时，已支持章节标题、列表、表格、常见内联样式，以及 EPUB 图片资源的嵌入。
- EPUB 正文标题翻译后会同步回写导航目录 `nav/ncx`，尽量保证封面、目录与正文标题一致，避免目录和章节名错位。
- 电子书导出 `DOCX` 时会生成封面页与可更新目录，目录基于译后标题层级生成。
- DOCX 封面页会优先复用 EPUB 封面图片及封面文档中的译后标题/副标题文本。
- DOCX 封面页会进一步带出 EPUB metadata 中的作者、出版社、语言、日期等书籍信息。
- EPUB 图片导出到 DOCX 时会优先读取资源真实像素尺寸，再结合 HTML/CSS 宽高信息进行缩放。
- EPUB 中的 `figure + img + figcaption` 会按图片块与图注组合导出，减少图片和图注分离。
- 电子书导出 `DOCX` 时会补出书籍信息/简介页，并支持 `SVG` 资源直通写入。
- 支持设置翻译范围：PDF 按页、PPT 按幻灯片、Excel 按工作表、EPUB 按章节文档、Word 按近似分页范围处理；结束值为 `0` 表示一直到末尾。
- `MOBI / AZW3` 导入时会先尝试转换为 `EPUB`；这一步需要本机 `Calibre` 的 `ebook-convert.exe`。
- 可通过 `publish.ps1` 打包为单文件 EXE。
- 可通过 `installer.iss` 与 `build-installer.ps1` 生成安装包。
- 提供两个辅助命令行工具：
  - `tools/PdfBilingualInspector`：检查 PDF 双语导出中的疑似未翻译片段
  - `tools/TranslatorCliRunner`：复用本机保存设置，在命令行直接重跑单个文档

## 运行

```powershell
dotnet run --project .\TranslatorApp\TranslatorApp.csproj
```

## 打包

```powershell
.\publish.ps1
```

打包结果默认输出到 `.\publish`。

如果你想要尽量纯单文件的独立版本：

```powershell
.\publish-singlefile.ps1
```

打包结果默认输出到 `.\publish-singlefile`。
这个版本会把 .NET 运行时打进 `TranslatorApp.exe`，不再额外携带 `appsettings.json`。
但为了避免把 OCR 语言数据塞进 exe，默认不会附带 `tessdata`；程序可以独立启动，如果要对扫描版 PDF 使用 OCR，请在界面中指定外部 `tessdata` 目录。

## OCR 准备

将 `eng.traineddata`、`chi_sim.traineddata` 放到 `.\tessdata`，或者在界面中指定 `tessdata` 目录。

## 电子书转换准备

如果要处理 `MOBI / AZW3`，请先安装 `Calibre`，并确保：

- `ebook-convert.exe` 已在系统 `PATH` 中，或
- 在软件“翻译设置”里手动指定 `ebook-convert.exe` 路径。

如果只处理 `EPUB`，则不依赖 `Calibre`。

## 术语表

不需要下载专门的术语表，直接使用项目里的示例文件即可：

- [glossary-template.txt](/e:/translator/glossary-template.txt)

你可以在这个文件基础上继续追加自己的术语。

## 安装包

安装 Inno Setup 后执行：

```powershell
.\build-installer.ps1
```

## 清理

类似 `make clean` 的本地清理脚本：

```powershell
.\clean.ps1
```

会删除：

- `TranslatorApp\bin`
- `TranslatorApp\obj`
- `bin_verify_inspector`
- `obj_cli_runner`
- `obj_verify_inspector`
- `publish`
- `installer-output`
- `tmp`

## 接手文档

为后续新的 Codex 会话准备的项目文档：

- [PROJECT_GUIDE.md](/e:/translator/PROJECT_GUIDE.md)
- [HANDOFF.md](/e:/translator/HANDOFF.md)
