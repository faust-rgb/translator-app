using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public sealed class TextDocumentTranslator(
    ITextTranslationService textTranslationService,
    IAppLogService logService,
    ITranslationProgressService progressService,
    IBilingualExportService bilingualExportService)
    : DocumentTranslatorBase(textTranslationService, logService)
{
    public override bool CanHandle(string extension) => extension == ".txt";

    public override async Task TranslateAsync(TranslationJobContext context)
    {
        var outputPath = BuildOutputPath(context.Item.SourcePath, context.Settings.Translation.OutputDirectory);
        context.Item.OutputPath = outputPath;

        var encoding = DetectTextEncoding(context.Item.SourcePath);
        var sourceText = await File.ReadAllTextAsync(context.Item.SourcePath, encoding, context.CancellationToken);
        var units = BuildTranslationUnits(sourceText);
        var requestedRange = GetRequestedRange(context.Settings, units.Count);
        var bilingualSegments = new List<BilingualSegment>();
        var outputBuilder = new StringBuilder(sourceText.Length + Math.Max(256, sourceText.Length / 4));

        for (var index = 0; index < units.Count; index++)
        {
            outputBuilder.Append(units[index].LeadingSeparator);

            var oneBasedIndex = index + 1;
            if (!IsWithinRequestedRange(oneBasedIndex, requestedRange))
            {
                outputBuilder.Append(units[index].OriginalText);
                continue;
            }

            if (index < context.ResumeUnitIndex)
            {
                outputBuilder.Append(units[index].OriginalText);
                continue;
            }

            var translated = await TranslateBlockAsync(
                units[index].OriginalText,
                units[index].ContextHint,
                "类型：纯文本段落。请保留段落语气、列表编号、缩进线索、URL、邮箱、命令行、代码片段和空行语义，不要擅自改写为其他文档格式。",
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            bilingualSegments.Add(new BilingualSegment(units[index].ContextHint, units[index].OriginalText, translated));
            outputBuilder.Append(translated);

            var processedInRange = Math.Min(oneBasedIndex, requestedRange.End) - requestedRange.Start + 1;
            var rangeLength = Math.Max(1, requestedRange.End - requestedRange.Start + 1);
            var progress = (int)Math.Round(processedInRange * 100d / rangeLength);
            await context.ReportProgressAsync(progress, $"TXT 文本块 {oneBasedIndex}/{units.Count}（范围 {requestedRange.Start}-{requestedRange.End}）");
            await context.SaveCheckpointAsync(oneBasedIndex, 0, $"TXT 文本块 {oneBasedIndex}/{units.Count}");
        }

        if (units.Count > 0)
        {
            outputBuilder.Append(units[^1].TrailingSeparator);
        }

        await File.WriteAllTextAsync(outputPath, outputBuilder.ToString(), encoding, context.CancellationToken);

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

    private static Encoding DetectTextEncoding(string path)
    {
        using var stream = File.OpenRead(path);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        _ = reader.Peek();
        return reader.CurrentEncoding;
    }

    private static List<TextTranslationUnit> BuildTranslationUnits(string sourceText)
    {
        var units = new List<TextTranslationUnit>();
        if (string.IsNullOrEmpty(sourceText))
        {
            return units;
        }

        var parts = Regex.Split(sourceText, @"(\r?\n\s*\r?\n+)");
        var pendingSeparator = string.Empty;
        var displayIndex = 0;

        foreach (var part in parts)
        {
            if (string.IsNullOrEmpty(part))
            {
                continue;
            }

            if (Regex.IsMatch(part, @"^\r?\n\s*\r?\n+$"))
            {
                pendingSeparator += part;
                continue;
            }

            if (string.IsNullOrWhiteSpace(part))
            {
                pendingSeparator += part;
                continue;
            }

            displayIndex++;
            units.Add(new TextTranslationUnit(
                pendingSeparator,
                part,
                string.Empty,
                $"TXT 文本块 {displayIndex}"));
            pendingSeparator = string.Empty;
        }

        if (units.Count == 0)
        {
            units.Add(new TextTranslationUnit(string.Empty, sourceText, string.Empty, "TXT 文本块 1"));
            return units;
        }

        units[^1] = units[^1] with { TrailingSeparator = pendingSeparator };
        return units;
    }

    private sealed record TextTranslationUnit(
        string LeadingSeparator,
        string OriginalText,
        string TrailingSeparator,
        string ContextHint);
}
