using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public sealed class WordDocumentTranslator(
    ITextTranslationService textTranslationService,
    IAppLogService logService,
    ITranslationProgressService progressService,
    IBilingualExportService bilingualExportService)
    : DocumentTranslatorBase(textTranslationService, logService)
{
    public override bool CanHandle(string extension) => extension == ".docx";

    public override async Task TranslateAsync(TranslationJobContext context)
    {
        var outputPath = BuildOutputPath(context.Item.SourcePath, context.Settings.Translation.OutputDirectory);
        if (!File.Exists(outputPath))
        {
            File.Copy(context.Item.SourcePath, outputPath, overwrite: true);
        }
        context.Item.OutputPath = outputPath;

        using var document = WordprocessingDocument.Open(outputPath, true);
        var bilingualSegments = new List<BilingualSegment>();

        var parts = new OpenXmlPartRootElement?[]
        {
            document.MainDocumentPart?.Document,
        }
        .Concat(document.MainDocumentPart?.HeaderParts.Select(x => x.Header) ?? Array.Empty<Header?>())
        .Concat(document.MainDocumentPart?.FooterParts.Select(x => x.Footer) ?? Array.Empty<Footer?>())
        .OfType<OpenXmlPartRootElement>()
        .ToList();

        var paragraphs = parts.SelectMany(x => x.Descendants<Paragraph>()).ToList();
        var batchSize = GetBlockTranslationConcurrency(context.Settings);
        for (var batchStart = context.ResumeUnitIndex; batchStart < paragraphs.Count; batchStart += batchSize)
        {
            var batchParagraphs = paragraphs
                .Skip(batchStart)
                .Take(batchSize)
                .Select((paragraph, offset) =>
                {
                    var runs = paragraph
                        .Descendants<Run>()
                        .Select(CreateRunInfo)
                        .Where(x => x is not null)
                        .Cast<WordRunInfo>()
                        .ToList();
                    var original = string.Concat(runs.Select(x => x.Original));
                    return new
                    {
                        ParagraphIndex = batchStart + offset,
                        Runs = runs,
                        Original = original
                    };
                })
                .Where(x => x.Runs.Count > 0 && !string.IsNullOrWhiteSpace(x.Original))
                .ToList();

            if (batchParagraphs.Count == 0)
            {
                continue;
            }

            var translatedBatch = await TranslateBatchAsync(
                batchParagraphs.Select(x => new TranslationBlock(x.Original, "Word 段落")).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var batchIndex = 0; batchIndex < batchParagraphs.Count; batchIndex++)
            {
                var paragraphInfo = batchParagraphs[batchIndex];
                var translated = translatedBatch[batchIndex];
                bilingualSegments.Add(new BilingualSegment("Word 段落", paragraphInfo.Original, translated));

                ApplyTranslationToParagraphRuns(paragraphInfo.Runs, translated);

                var progress = (int)Math.Round((paragraphInfo.ParagraphIndex + 1) * 100d / Math.Max(1, paragraphs.Count));
                await context.ReportProgressAsync(progress, $"Word 段落 {paragraphInfo.ParagraphIndex + 1}/{paragraphs.Count}");
                await context.SaveCheckpointAsync(paragraphInfo.ParagraphIndex + 1, 0, $"Word 段落 {paragraphInfo.ParagraphIndex + 1}/{paragraphs.Count}");
            }
        }

        foreach (var part in parts)
        {
            part.Save();
        }

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static void ApplyTranslationToParagraphRuns(IReadOnlyList<WordRunInfo> runs, string translated)
    {
        if (runs.Count == 0)
        {
            return;
        }

        if (runs.Count == 1)
        {
            ApplySegmentToTexts(runs[0].Texts, translated);
            return;
        }

        var formatGroups = GroupRunsByFormat(runs);
        if (formatGroups.Count == 1)
        {
            ApplySegmentToTexts(formatGroups[0].Texts, translated);
            return;
        }

        var segments = TextDistributionHelper.Distribute(
            translated,
            formatGroups.Select(x => Math.Max(1, x.Original.Length)).ToList());

        for (var i = 0; i < formatGroups.Count; i++)
        {
            ApplySegmentToTexts(formatGroups[i].Texts, segments[i]);
        }
    }

    private static List<WordFormatGroup> GroupRunsByFormat(IReadOnlyList<WordRunInfo> runs)
    {
        var groups = new List<WordFormatGroup>();

        foreach (var run in runs)
        {
            if (groups.Count > 0 && string.Equals(groups[^1].FormatKey, run.FormatKey, StringComparison.Ordinal))
            {
                groups[^1].Texts.AddRange(run.Texts);
                groups[^1].Original += run.Original;
                continue;
            }

            groups.Add(new WordFormatGroup(run.FormatKey, run.Texts.ToList(), run.Original));
        }
        
        return groups;
    }

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial)
    {
        progressService.Publish(Path.GetFileName(sourcePath), partial);
        return Task.CompletedTask;
    }

    private static WordRunInfo? CreateRunInfo(Run run)
    {
        var texts = run.Elements<Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
        if (texts.Count == 0)
        {
            return null;
        }

        return new WordRunInfo(texts, string.Concat(texts.Select(x => x.Text)), GetFormatKey(run));
    }

    private static string GetFormatKey(Run run)
    {
        var properties = run.RunProperties;
        return properties?.OuterXml ?? string.Empty;
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

        // 清空其他 Text 节点
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

    private sealed class WordFormatGroup(string formatKey, List<Text> texts, string original)
    {
        public string FormatKey { get; } = formatKey;
        public List<Text> Texts { get; } = texts;
        public string Original { get; set; } = original;
    }
    private sealed record WordRunInfo(IReadOnlyList<Text> Texts, string Original, string FormatKey);
}
