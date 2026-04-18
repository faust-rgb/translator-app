# Handoff

## What Is Already Done

- Local WPF app built from scratch in `e:\translator`
- Multi-format translation for `docx/xlsx/pptx/pdf/epub/mobi/azw3`
- OCR fallback for scanned PDFs
- Sequential queue behavior
- Task resume / delete / open output actions
- Connection test
- History persistence
- GitHub repo created and pushed
- PDF translation heuristics upgraded:
  - custom PDFsharp font resolver for CJK output
  - paragraph-style block merging instead of line-by-line translation
  - cross-block hyphenation and continuation repair before translation
  - marginal `arXiv` side-note filtering
  - character-based wrapping for Chinese PDF output
  - pure formula blocks are preserved, while prose containing formulas is still translated with formula-preservation hints
- PDF rendering heuristics further upgraded:
  - OCR blocks are filtered for likely page-number / header-footer noise before redraw
  - blocks are classified as title / caption / header-footer / footnote / list / code / table row
  - redraw uses block-aware margin, line-height, hanging-indent, and overflow fallback strategies
  - OCR sparse-text detection, OCR column/block aggregation, PDF column detection, marginal-noise filtering, and paragraph-continuation thresholds are now configurable through settings instead of being hard-coded in one place
- PDF text translation now retries suspicious English fragments with stronger instructions, and the retry path creates a fresh AI client per attempt to avoid stale-request failures after 504s.
- Added CLI helpers:
  - `tools/PdfBilingualInspector` for bilingual export inspection
  - `tools/TranslatorCliRunner` for headless single-document reruns using saved local settings
- Word translation preserves formatting more safely by redistributing translated text at continuous-format-group boundaries instead of raw run-length splits
- Word range translation now uses approximate source-page detection from rendered/manual page breaks plus section/page-break anchors
- Translation concurrency is now explicitly split into:
  - document-level parallelism
  - block-level parallelism
  - global remote-request throttling
  Default safe mode is `1/1/1`
- Stream preview panel now wraps text to the visible width
- Task list context menu now selects the right-clicked row before executing actions
- Native ebook pipeline added:
  - EPUB can be translated and exported natively as EPUB or DOCX
  - MOBI/AZW3 import through `ebook-convert.exe` to EPUB first
  - EPUB body headings sync back into `nav/ncx` TOC labels
  - DOCX ebook export supports cover page, metadata/info page, updateable TOC, images, figures, captions, and common inline/block styles
  - figure/image export now reads real pixel dimensions when available
- Translation range setting added in UI and settings:
  - PDF = pages
  - PPT = slides
  - Excel = worksheets
  - EPUB = chapter/content documents
  - Word = approximate source pages

## Repository

- `https://github.com/faust-rgb/translator-app`

## Common Commands

```powershell
dotnet build .\TranslatorApp\TranslatorApp.csproj
dotnet run --project .\TranslatorApp\TranslatorApp.csproj
.\publish.ps1
.\clean.ps1
```

## Known Operational Caveats

- `publish.ps1` cannot overwrite `publish\TranslatorApp.exe` while the app is still open.
- OCR requires `tessdata` files such as `eng.traineddata` and `chi_sim.traineddata`.
- Third-party model gateways are only "compatible", so endpoint quirks are expected.
- PDF resume checkpoints are still saved, but PDF output currently regenerates from page 1 after restart because PDFsharp documents cannot be incrementally re-saved and then modified again.
- PDF layout is much improved, but academic PDFs with dense figures/equations can still need heuristic tuning in `PdfDocumentTranslator.cs`.
- Most practical PDF/OCR tuning knobs are now surfaced in `TranslationSettings` / `OcrSettings`, so adjust config first before changing code.
- Inline formulas inside normal prose are now handled more carefully, but formula/prose boundary heuristics are still best-effort and should be regression-checked against papers with dense notation.
- PDF tables and charts are still handled heuristically: original shapes/images stay in place, while detectable text is translated and redrawn.
- Word "page range" is still approximate rather than Word-renderer-perfect.
- EPUB "range" is chapter/content-document based, not reader page-number based.
- Partial EPUB translation intentionally avoids fully re-translating untouched TOC entries, so mixed translated/untranslated navigation is possible when only some chapters are selected.

## Where To Edit Next

- UI layout: `TranslatorApp/MainWindow.xaml`
- AI protocol behavior: `TranslatorApp/Services/Ai`
- Recovery flow: `TranslatorApp/Services/RecoveryStateService.cs`
- PDF/OCR behavior: `TranslatorApp/Services/Documents/PdfDocumentTranslator.cs`, `TranslatorApp/Services/OcrService.cs`
- Ebook behavior: `TranslatorApp/Services/Documents/EbookDocumentTranslator.cs`, `TranslatorApp/Services/Documents/EbookDocxExportService.cs`
- Remote request throttling: `TranslatorApp/Services/TextTranslationService.cs`, `TranslatorApp/Services/TranslationRequestThrottle.cs`
