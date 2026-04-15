using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public sealed class ExcelDocumentTranslator(
    ITextTranslationService textTranslationService,
    IAppLogService logService,
    ITranslationProgressService progressService,
    IBilingualExportService bilingualExportService)
    : DocumentTranslatorBase(textTranslationService, logService)
{
    public override bool CanHandle(string extension) => extension == ".xlsx";

    public override async Task TranslateAsync(TranslationJobContext context)
    {
        var outputPath = BuildOutputPath(context.Item.SourcePath, context.Settings.Translation.OutputDirectory);
        if (!File.Exists(outputPath))
        {
            File.Copy(context.Item.SourcePath, outputPath, overwrite: true);
        }
        context.Item.OutputPath = outputPath;

        using var document = SpreadsheetDocument.Open(outputPath, true);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("Excel 文件无效。");
        var bilingualSegments = new List<BilingualSegment>();

        var actions = new List<Func<Task>>();

        if (workbookPart.SharedStringTablePart?.SharedStringTable is { } sharedTable)
        {
            foreach (var item in sharedTable.Elements<SharedStringItem>())
            {
                var texts = item.Descendants<Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
                if (texts.Count == 0)
                {
                    continue;
                }

                actions.Add(async () =>
                {
                    var original = string.Concat(texts.Select(x => x.Text));
                    var translated = await TranslateBlockAsync(
                        original,
                        "Excel 共享字符串",
                        context.Settings,
                        context.PauseController,
                        partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                        context.CancellationToken);
                    bilingualSegments.Add(new BilingualSegment("Excel 共享字符串", original, translated));

                    var segments = TextDistributionHelper.Distribute(translated, texts.Select(x => Math.Max(1, x.Text?.Length ?? 0)).ToList());
                    for (var i = 0; i < texts.Count; i++)
                    {
                        texts[i].Space = SpaceProcessingModeValues.Preserve;
                        texts[i].Text = segments[i];
                    }
                });
            }
        }

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet;
            if (worksheet is null)
            {
                continue;
            }

            foreach (var cell in worksheet.Descendants<Cell>().Where(x => x.DataType?.Value == CellValues.InlineString))
            {
                var texts = cell.Descendants<Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
                if (texts.Count == 0)
                {
                    continue;
                }

                actions.Add(async () =>
                {
                    var original = string.Concat(texts.Select(x => x.Text));
                    var translated = await TranslateBlockAsync(
                        original,
                        "Excel 行内文本",
                        context.Settings,
                        context.PauseController,
                        partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                        context.CancellationToken);
                    bilingualSegments.Add(new BilingualSegment("Excel 行内文本", original, translated));

                    var segments = TextDistributionHelper.Distribute(translated, texts.Select(x => Math.Max(1, x.Text?.Length ?? 0)).ToList());
                    for (var i = 0; i < texts.Count; i++)
                    {
                        texts[i].Space = SpaceProcessingModeValues.Preserve;
                        texts[i].Text = segments[i];
                    }
                });
            }
        }

        for (var index = 0; index < actions.Count; index++)
        {
            if (index < context.ResumeUnitIndex)
            {
                continue;
            }

            await actions[index]();
            var progress = (int)Math.Round((index + 1) * 100d / Math.Max(1, actions.Count));
            await context.ReportProgressAsync(progress, $"Excel 单元 {index + 1}/{actions.Count}");
            await context.SaveCheckpointAsync(index + 1, 0, $"Excel 单元 {index + 1}/{actions.Count}");
            workbookPart.Workbook?.Save();
        }

        workbookPart.Workbook?.Save();
        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial) =>
        Task.Run(() => progressService.Publish(Path.GetFileName(sourcePath), partial));
}
