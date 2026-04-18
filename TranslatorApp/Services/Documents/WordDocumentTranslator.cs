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
        var mainDocument = document.MainDocumentPart?.Document ?? throw new InvalidOperationException("Word 文件无效。");
        var mainParagraphs = mainDocument.Body?.Descendants<Paragraph>().ToList() ?? [];
        var pageInfos = BuildWordPageInfos(mainParagraphs);
        var totalPages = Math.Max(1, pageInfos.LastOrDefault()?.PageNumber ?? 1);
        var requestedRange = GetRequestedRange(context.Settings, totalPages);
        var paragraphs = pageInfos
            .Where(x => IsWithinRequestedRange(x.PageNumber, requestedRange))
            .Select(x => x.Paragraph)
            .ToList();

        var translateHeadersAndFooters = requestedRange.Start == 1 && requestedRange.End >= totalPages;
        var parts = new List<OpenXmlPartRootElement> { mainDocument };
        if (translateHeadersAndFooters)
        {
            parts.AddRange(document.MainDocumentPart?.HeaderParts.Select(x => x.Header).OfType<Header>() ?? []);
            parts.AddRange(document.MainDocumentPart?.FooterParts.Select(x => x.Footer).OfType<Footer>() ?? []);
            paragraphs.AddRange(parts.Skip(1).SelectMany(x => x.Descendants<Paragraph>()));
        }
        else if (document.MainDocumentPart?.HeaderParts.Any() == true || document.MainDocumentPart?.FooterParts.Any() == true)
        {
            Log($"Word 当前按近似分页范围 {requestedRange.Start}-{requestedRange.End} 处理，页眉页脚仅在全量范围时一并翻译。");
        }

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
                await context.ReportProgressAsync(progress, $"Word 段落 {paragraphInfo.ParagraphIndex + 1}/{paragraphs.Count}（近似页范围 {requestedRange.Start}-{requestedRange.End}）");
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

    private static List<WordPageInfo> BuildWordPageInfos(IReadOnlyList<Paragraph> paragraphs)
    {
        var infos = new List<WordPageInfo>(paragraphs.Count);
        var currentPage = 1;
        foreach (var paragraph in paragraphs)
        {
            if (infos.Count > 0 && StartsNewPage(paragraph))
            {
                currentPage++;
            }

            infos.Add(new WordPageInfo(paragraph, currentPage));

            var renderedBreaks = paragraph.Descendants<LastRenderedPageBreak>().Count();
            var manualBreaks = paragraph
                .Descendants<Break>()
                .Count(x => x.Type?.Value == BreakValues.Page);
            currentPage += Math.Max(renderedBreaks, 0) + Math.Max(manualBreaks, 0);
        }

        return infos;
    }

    private static bool StartsNewPage(Paragraph paragraph)
    {
        if (paragraph.ParagraphProperties?.PageBreakBefore is not null)
        {
            return true;
        }

        var sectionType = paragraph.ParagraphProperties?
            .SectionProperties?
            .GetFirstChild<SectionType>()?
            .Val?
            .Value;

        return sectionType == SectionMarkValues.NextPage ||
               sectionType == SectionMarkValues.OddPage ||
               sectionType == SectionMarkValues.EvenPage;
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
    private sealed record WordPageInfo(Paragraph Paragraph, int PageNumber);
}
