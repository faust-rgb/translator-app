using TranslatorApp.Configuration;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public interface IOcrService
{
    Task<IReadOnlyList<OcrTextBlock>> RecognizePdfPageAsync(string pdfPath, int pageIndex, OcrSettings settings, CancellationToken cancellationToken);
}
