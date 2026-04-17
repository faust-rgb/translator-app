using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class BilingualExportService(IAppLogService logService) : IBilingualExportService
{
    public Task ExportAsync(string sourcePath, string outputDirectory, IReadOnlyList<BilingualSegment> segments, CancellationToken cancellationToken)
    {
        if (segments.Count == 0)
        {
            return Task.CompletedTask;
        }

        logService.Info("正在导出双语对照文档。若原文包含敏感内容，导出文件会同时包含原文和译文。");

        var baseDirectory = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.GetDirectoryName(sourcePath)!
            : outputDirectory;
        Directory.CreateDirectory(baseDirectory);

        var bilingualPath = Path.Combine(
            baseDirectory,
            $"{Path.GetFileNameWithoutExtension(sourcePath)}.bilingual.docx");

        using var document = WordprocessingDocument.Create(bilingualPath, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        body.AppendChild(new Paragraph(new Run(new Text("双语对照导出"))));

        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        table.Append(CreateRow("上下文", "原文", "译文", isHeader: true));
        foreach (var segment in segments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            table.Append(CreateRow(segment.ContextHint, segment.Original, segment.Translation, isHeader: false));
        }

        body.Append(table);
        mainPart.Document.Save();
        return Task.CompletedTask;
    }

    private static TableRow CreateRow(string context, string original, string translation, bool isHeader)
    {
        var row = new TableRow();
        row.Append(CreateCell(context, isHeader));
        row.Append(CreateCell(original, isHeader));
        row.Append(CreateCell(translation, isHeader));
        return row;
    }

    private static TableCell CreateCell(string text, bool isHeader)
    {
        var runProperties = new RunProperties();
        if (isHeader)
        {
            runProperties.Bold = new Bold();
        }

        var run = new Run(runProperties, new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        var paragraph = new Paragraph(run);
        return new TableCell(paragraph);
    }
}
