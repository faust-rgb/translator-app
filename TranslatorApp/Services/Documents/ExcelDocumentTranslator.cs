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

        var items = new List<ExcelTranslationItem>();

        if (workbookPart.SharedStringTablePart?.SharedStringTable is { } sharedTable)
        {
            foreach (var item in sharedTable.Elements<SharedStringItem>())
            {
                var runs = item
                    .Elements<Run>()
                    .Select(CreateRunInfo)
                    .Where(x => x is not null)
                    .Cast<ExcelRunInfo>()
                    .ToList();
                if (runs.Count == 0)
                {
                    continue;
                }

                items.Add(new ExcelTranslationItem("Excel 共享字符串", string.Concat(runs.Select(x => x.Original)), runs));
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
                var runs = cell
                    .Descendants<Run>()
                    .Select(CreateRunInfo)
                    .Where(x => x is not null)
                    .Cast<ExcelRunInfo>()
                    .ToList();
                if (runs.Count == 0)
                {
                    continue;
                }

                items.Add(new ExcelTranslationItem("Excel 行内文本", string.Concat(runs.Select(x => x.Original)), runs));
            }
        }

        var batchSize = GetBlockTranslationConcurrency(context.Settings);
        for (var batchStart = context.ResumeUnitIndex; batchStart < items.Count; batchStart += batchSize)
        {
            var batch = items.Skip(batchStart).Take(batchSize).ToList();
            if (batch.Count == 0)
            {
                continue;
            }

            var translatedBatch = await TranslateBatchAsync(
                batch.Select(x => new TranslationBlock(x.Original, x.ContextHint)).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var batchIndex = 0; batchIndex < batch.Count; batchIndex++)
            {
                var item = batch[batchIndex];
                var translated = translatedBatch[batchIndex];
                bilingualSegments.Add(new BilingualSegment(item.ContextHint, item.Original, translated));

                var segments = TextDistributionHelper.Distribute(translated, item.Runs.Select(x => Math.Max(1, x.Original.Length)).ToList());
                for (var i = 0; i < item.Runs.Count; i++)
                {
                    ApplySegmentToTexts(item.Runs[i].Texts, segments[i]);
                }

                var absoluteIndex = batchStart + batchIndex;
                var progress = (int)Math.Round((absoluteIndex + 1) * 100d / Math.Max(1, items.Count));
                await context.ReportProgressAsync(progress, $"Excel 单元 {absoluteIndex + 1}/{items.Count}");
                await context.SaveCheckpointAsync(absoluteIndex + 1, 0, $"Excel 单元 {absoluteIndex + 1}/{items.Count}");
            }

            workbookPart.Workbook?.Save();
        }

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial)
    {
        progressService.Publish(Path.GetFileName(sourcePath), partial);
        return Task.CompletedTask;
    }

    private static ExcelRunInfo? CreateRunInfo(Run run)
    {
        var texts = run.Elements<Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
        if (texts.Count == 0)
        {
            return null;
        }

        return new ExcelRunInfo(texts, string.Concat(texts.Select(x => x.Text)));
    }

    private static void ApplySegmentToTexts(IReadOnlyList<Text> texts, string segment)
    {
        if (texts.Count == 0)
        {
            return;
        }

        // 保留原始首尾空格信息
        var originalFirstText = texts[0].Text ?? string.Empty;
        var leadingSpace = GetLeadingWhitespace(originalFirstText);
        var trailingSpace = GetTrailingWhitespace(texts[^1].Text ?? string.Empty);

        // 应用译文，保留空格
        var processedSegment = leadingSpace + segment.Trim() + trailingSpace;

        texts[0].Space = SpaceProcessingModeValues.Preserve;
        texts[0].Text = processedSegment;

        for (var i = 1; i < texts.Count; i++)
        {
            texts[i].Space = SpaceProcessingModeValues.Preserve;
            texts[i].Text = string.Empty;
        }
    }

    private static string GetLeadingWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var index = 0;
        while (index < text.Length && char.IsWhiteSpace(text[index]))
        {
            index++;
        }

        return text[..index];
    }

    private static string GetTrailingWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var index = text.Length - 1;
        while (index >= 0 && char.IsWhiteSpace(text[index]))
        {
            index--;
        }

        return text[(index + 1)..];
    }

    private sealed record ExcelTranslationItem(string ContextHint, string Original, IReadOnlyList<ExcelRunInfo> Runs);
    private sealed record ExcelRunInfo(IReadOnlyList<Text> Texts, string Original);
}
