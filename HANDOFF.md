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

## Where To Edit Next

- UI layout: `TranslatorApp/MainWindow.xaml`
- AI protocol behavior: `TranslatorApp/Services/Ai`
- Recovery flow: `TranslatorApp/Services/RecoveryStateService.cs`
- PDF/OCR behavior: `TranslatorApp/Services/Documents/PdfDocumentTranslator.cs`, `TranslatorApp/Services/OcrService.cs`
