# Project Guide

## Purpose

`TranslatorApp` is a local Windows desktop translator for Office, PDF, and ebook documents.  
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
- `WIC / BitmapDecoder` for ebook image sizing during DOCX export

## Main Structure

- `TranslatorApp/ViewModels/MainViewModel.cs`
  Main UI state, commands, startup restore, history refresh, connection test, translation range settings.
- `TranslatorApp/MainWindow.xaml`
  Main desktop layout.
- `TranslatorApp/Services/Ai`
  Provider clients and endpoint resolution.
- `TranslatorApp/Services/TranslationRequestThrottle.cs`
  Global remote translation request throttling.
- `TranslatorApp/Services/Documents`
  Translators for `docx/xlsx/pptx/pdf/epub/mobi/azw3`.
- `TranslatorApp/Services/Documents/EbookDocumentTranslator.cs`
  Native EPUB translation pipeline, TOC synchronization, cover/metadata extraction, partial chapter-range support.
- `TranslatorApp/Services/Documents/EbookDocxExportService.cs`
  Native EPUB-to-DOCX export with cover page, metadata page, TOC field, images, figures, captions, and common inline/block style mapping.
- `TranslatorApp/Services/OcrService.cs`
  OCR fallback for scanned PDFs.
- `tools/PdfBilingualInspector`
  Inspect bilingual PDF exports for likely untranslated English fragments.
- `tools/TranslatorCliRunner`
  Re-run a saved document translation from the command line using local user settings.
- `TranslatorApp/Services/RecoveryStateService.cs`
  Crash recovery and resume state.
- `TranslatorApp/Services/TranslationHistoryService.cs`
  Persistent run history.

## Current UX Rules

- User config is intentionally simplified for third-party compatible endpoints.
- Documents run sequentially by default.
- Translation range is configurable:
  - PDF uses source pages
  - PowerPoint uses source slides
  - Excel uses source worksheets
  - EPUB/MOBI/AZW3 use chapter/content-document ranges
  - Word uses approximate source-page ranges based on page-break anchors
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
  - repairs some cross-block hyphenation / continuation splits before translation
  - distinguishes pure formula blocks from prose that merely contains formulas
  - classifies likely titles, captions, header/footer lines, footnotes, lists, code, and table-like rows
  - uses block-aware redraw margins and overflow fallback
  - wraps translated text character-by-character for Chinese output
- PDF retry logic now does an extra pass for risky English fragments that come back untranslated, and creates a fresh AI client for each retry attempt.
- Word translation no longer redistributes translated text at raw run granularity; it groups adjacent runs with identical formatting first to preserve formatting boundaries more safely.
- Ebook support now follows this model:
  - native EPUB translation edits XHTML/TOC locally
  - EPUB output is native
  - DOCX output is native and keeps book structure as closely as practical
  - MOBI/AZW3 are imported through `ebook-convert.exe` into EPUB first
- EPUB TOC synchronization now prefers translated body headings so navigation labels stay aligned with chapter titles.
- DOCX ebook export now includes:
  - cover page from EPUB cover image and cover-document text
  - metadata/info page from EPUB metadata
  - update-on-open TOC field
  - image embedding with real pixel-size detection when possible
  - `figure + figcaption` grouped export
  - SVG passthrough via OpenXML image-part support
- If only part of an EPUB is translated, TOC synchronization is limited to translated chapters to avoid mixing rewritten and untouched navigation labels.
- If `publish` fails with access denied, the existing `publish\TranslatorApp.exe` is usually still running.
