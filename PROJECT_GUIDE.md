# Project Guide

## Purpose

`TranslatorApp` is a local Windows desktop translator for Office and PDF documents.  
The app keeps document processing local and only sends extracted text to an AI model endpoint.

## Tech Stack

- `.NET 10`
- `WPF`
- `DocumentFormat.OpenXml`
- `PDFsharp`
- `UglyToad.PdfPig`
- `Docnet.Core`
- `Tesseract`
- `CommunityToolkit.Mvvm`

## Main Structure

- `TranslatorApp/ViewModels/MainViewModel.cs`
  Main UI state, commands, startup restore, history refresh, connection test.
- `TranslatorApp/MainWindow.xaml`
  Main desktop layout.
- `TranslatorApp/Services/Ai`
  Provider clients and endpoint resolution.
- `TranslatorApp/Services/TranslationRequestThrottle.cs`
  Global remote translation request throttling.
- `TranslatorApp/Services/Documents`
  Translators for `docx/xlsx/pptx/pdf`.
- `TranslatorApp/Services/OcrService.cs`
  OCR fallback for scanned PDFs.
- `TranslatorApp/Services/RecoveryStateService.cs`
  Crash recovery and resume state.
- `TranslatorApp/Services/TranslationHistoryService.cs`
  Persistent run history.

## Current UX Rules

- User config is intentionally simplified for third-party compatible endpoints.
- Documents run sequentially by default.
- Safe remote-request mode is enabled by default:
  - document parallelism = 1
  - block parallelism = 1
  - global remote request concurrency = 1
- Stopped tasks do not auto-restart.
- Resume is explicit from the task context menu.

## Run

```powershell
dotnet run --project .\TranslatorApp\TranslatorApp.csproj
```

## Build

```powershell
.\publish.ps1
```

## Clean

```powershell
.\clean.ps1
```

## Notes For Next Codex Session

- OpenAI-compatible endpoints may need `chat/completions` under `/v1`, `/v2`, or `/v3`.
- Anthropic-compatible endpoints may require different auth header styles; client already tries several.
- Settings import/export is supported from the main window.
- UI layout has been tuned for smaller displays, but further pixel-level adjustment should be done against screenshots.
- Task list context-menu actions are wired in `MainWindow.xaml.cs` to avoid WPF `ContextMenu` binding edge cases.
- PDF translation currently relies on heuristic block reconstruction:
  - filters marginal noise such as `arXiv` sidebars
  - merges nearby lines into paragraph-like blocks
  - preserves formula-like blocks instead of translating them
  - classifies likely titles, captions, header/footer lines, footnotes, lists, code, and table-like rows
  - uses block-aware redraw margins and overflow fallback
  - wraps translated text character-by-character for Chinese output
- Word translation no longer redistributes translated text at raw run granularity; it groups adjacent runs with identical formatting first to preserve formatting boundaries more safely.
- If `publish` fails with access denied, the existing `publish\TranslatorApp.exe` is usually still running.
