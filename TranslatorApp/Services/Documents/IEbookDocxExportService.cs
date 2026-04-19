namespace TranslatorApp.Services.Documents;

public interface IEbookDocxExportService
{
    Task ExportAsync(
        string outputPath,
        string title,
        EbookDocumentTranslator.EpubCoverInfo? cover,
        EbookDocumentTranslator.EpubMetadata metadata,
        IReadOnlyList<EbookDocumentTranslator.EpubExportDocument> contentDocuments,
        CancellationToken cancellationToken);
}
