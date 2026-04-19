using System.IO;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public sealed class PowerPointDocumentTranslator(
    ITextTranslationService textTranslationService,
    IAppLogService logService,
    ITranslationProgressService progressService,
    IBilingualExportService bilingualExportService)
    : DocumentTranslatorBase(textTranslationService, logService)
{
    public override bool CanHandle(string extension) => extension == ".pptx";

    public override async Task TranslateAsync(TranslationJobContext context)
    {
        var outputPath = BuildOutputPath(context.Item.SourcePath, context.Settings.Translation.OutputDirectory);
        if (!File.Exists(outputPath))
        {
            File.Copy(context.Item.SourcePath, outputPath, overwrite: true);
        }
        context.Item.OutputPath = outputPath;

        using var presentation = PresentationDocument.Open(outputPath, true);
        var slideParts = presentation.PresentationPart?.SlideParts.ToList() ?? new List<SlidePart>();
        var bilingualSegments = new List<BilingualSegment>();
        var requestedRange = GetRequestedRange(context.Settings, slideParts.Count);
        var selectedSlideCount = Math.Max(0, requestedRange.End - requestedRange.Start + 1);
        var processedSlides = 0;

        for (var slideIndex = 0; slideIndex < slideParts.Count; slideIndex++)
        {
            var slideNumber = slideIndex + 1;
            if (!IsWithinRequestedRange(slideNumber, requestedRange))
            {
                continue;
            }

            if (processedSlides < context.ResumeUnitIndex)
            {
                processedSlides++;
                continue;
            }

            var slide = slideParts[slideIndex].Slide;
            if (slide is null)
            {
                continue;
            }

            var paragraphs = slide.Descendants<A.Paragraph>().ToList();
            var paragraphInfos = paragraphs
                .Select(paragraph =>
                {
                    var runs = paragraph
                        .Descendants<A.Run>()
                        .Select(CreateRunInfo)
                        .Where(x => x is not null)
                        .Cast<PowerPointRunInfo>()
                        .ToList();
                    var original = string.Concat(runs.Select(x => x.Original));
                    return new
                    {
                        Runs = runs,
                        Original = original
                    };
                })
                .Where(x => x.Runs.Count > 0 && !string.IsNullOrWhiteSpace(x.Original))
                .ToList();

            var batchSize = GetBlockTranslationConcurrency(context.Settings);
            for (var batchStart = 0; batchStart < paragraphInfos.Count; batchStart += batchSize)
            {
                var batch = paragraphInfos.Skip(batchStart).Take(batchSize).ToList();
                var translatedBatch = await TranslateBatchAsync(
                    batch.Select(x => new TranslationBlock(x.Original, "PowerPoint 文本框段落")).ToList(),
                    context.Settings,
                    context.PauseController,
                    partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                    context.CancellationToken);

                for (var batchIndex = 0; batchIndex < batch.Count; batchIndex++)
                {
                    var paragraphInfo = batch[batchIndex];
                    var translated = translatedBatch[batchIndex];
                    bilingualSegments.Add(new BilingualSegment("PowerPoint 文本框段落", paragraphInfo.Original, translated));

                    var formatGroups = FormattedTextRunHelper.GroupAdjacentRunsByFormat(
                        paragraphInfo.Runs.Select(run => new FormattedTextRun<A.Text>(run.Texts, run.Original, run.FormatKey)).ToList());
                    var segments = FormattedTextRunHelper.DistributeAcrossGroups(translated, formatGroups);
                    for (var i = 0; i < formatGroups.Count; i++)
                    {
                        ApplySegmentToTexts(formatGroups[i].Texts, segments[i]);
                    }
                }
            }

            slide.Save();
            processedSlides++;
            var progress = (int)Math.Round(processedSlides * 100d / Math.Max(1, selectedSlideCount));
            await context.ReportProgressAsync(progress, $"PPT 幻灯片 {slideNumber}/{slideParts.Count}（范围 {requestedRange.Start}-{requestedRange.End}）");
            await context.SaveCheckpointAsync(processedSlides, 0, $"PPT 幻灯片 {slideNumber}/{slideParts.Count}");
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

    private static PowerPointRunInfo? CreateRunInfo(A.Run run)
    {
        var texts = run.Elements<A.Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
        if (texts.Count == 0)
        {
            return null;
        }

        return new PowerPointRunInfo(
            texts,
            string.Concat(texts.Select(x => x.Text)),
            GetFormatKey(run));
    }

    private static string GetFormatKey(A.Run run) =>
        run.GetFirstChild<A.RunProperties>()?.OuterXml ?? string.Empty;

    private static readonly OpenXmlAttribute XmlSpacePreserve =
        new("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");

    private static void ApplySegmentToTexts(IReadOnlyList<A.Text> texts, string segment)
    {
        if (texts.Count == 0)
        {
            return;
        }

        // 保留原始首尾空格信息
        var originalFirstText = texts[0].Text ?? string.Empty;
        var leadingSpace = WhitespacePreservationHelper.GetLeadingWhitespace(originalFirstText);
        var trailingSpace = WhitespacePreservationHelper.GetTrailingWhitespace(texts[^1].Text ?? string.Empty);

        // 应用译文，保留空格
        var processedSegment = leadingSpace + segment.Trim() + trailingSpace;

        texts[0].SetAttribute(XmlSpacePreserve);
        texts[0].Text = processedSegment;

        for (var i = 1; i < texts.Count; i++)
        {
            texts[i].SetAttribute(XmlSpacePreserve);
            texts[i].Text = string.Empty;
        }
    }

    private sealed record PowerPointRunInfo(IReadOnlyList<A.Text> Texts, string Original, string FormatKey);
}
