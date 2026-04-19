namespace TranslatorApp.Services.Documents;

public interface IEbookConversionService
{
    Task ConvertAsync(string inputPath, string outputPath, string configuredExecutablePath, CancellationToken cancellationToken);
}
