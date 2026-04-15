using System.IO;
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

        for (var slideIndex = 0; slideIndex < slideParts.Count; slideIndex++)
        {
            if (slideIndex < context.ResumeUnitIndex)
            {
                continue;
            }

            var slide = slideParts[slideIndex].Slide;
            if (slide is null)
            {
                continue;
            }

            var paragraphs = slide.Descendants<A.Paragraph>().ToList();
            foreach (var paragraph in paragraphs)
            {
                var texts = paragraph.Descendants<A.Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
                if (texts.Count == 0)
                {
                    continue;
                }

                var original = string.Concat(texts.Select(x => x.Text));
                var translated = await TranslateBlockAsync(
                    original,
                    "PowerPoint 文本框段落",
                    context.Settings,
                    context.PauseController,
                    partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                    context.CancellationToken);
                bilingualSegments.Add(new BilingualSegment("PowerPoint 文本框段落", original, translated));

                var segments = TextDistributionHelper.Distribute(translated, texts.Select(x => Math.Max(1, x.Text?.Length ?? 0)).ToList());
                for (var i = 0; i < texts.Count; i++)
                {
                    texts[i].Text = segments[i];
                }
            }

            slide.Save();
            var progress = (int)Math.Round((slideIndex + 1) * 100d / Math.Max(1, slideParts.Count));
            await context.ReportProgressAsync(progress, $"PPT 幻灯片 {slideIndex + 1}/{slideParts.Count}");
            await context.SaveCheckpointAsync(slideIndex + 1, 0, $"PPT 幻灯片 {slideIndex + 1}/{slideParts.Count}");
        }

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial) =>
        Task.Run(() => progressService.Publish(Path.GetFileName(sourcePath), partial));
}
