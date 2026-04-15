# Handoff

## What Is Already Done

- Local WPF app built from scratch in `e:\translator`
- Multi-format translation for `docx/xlsx/pptx/pdf`
- OCR fallback for scanned PDFs
- Sequential queue behavior
- Task resume / delete / open output actions
- Connection test
- History persistence
- GitHub repo created and pushed
- PDF translation heuristics upgraded:
  - custom PDFsharp font resolver for CJK output
  - paragraph-style block merging instead of line-by-line translation
  - marginal `arXiv` side-note filtering
  - character-based wrapping for Chinese PDF output
  - formula-like blocks are preserved instead of translated/overwritten
- Task list context menu now selects the right-clicked row before executing actions

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

## Where To Edit Next

- UI layout: `TranslatorApp/MainWindow.xaml`
- AI protocol behavior: `TranslatorApp/Services/Ai`
- Recovery flow: `TranslatorApp/Services/RecoveryStateService.cs`
- PDF/OCR behavior: `TranslatorApp/Services/Documents/PdfDocumentTranslator.cs`, `TranslatorApp/Services/OcrService.cs`
