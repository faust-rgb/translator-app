using TranslatorApp.Models;

namespace TranslatorApp.Services;

public interface IBilingualExportService
{
    Task ExportAsync(string sourcePath, string outputDirectory, IReadOnlyList<BilingualSegment> segments, CancellationToken cancellationToken);
}
