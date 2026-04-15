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
        for (var index = 0; index < paragraphs.Count; index++)
        {
            if (index < context.ResumeUnitIndex)
            {
                continue;
            }

            var paragraph = paragraphs[index];
            var texts = paragraph.Descendants<Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
            var original = string.Concat(texts.Select(x => x.Text));
            if (string.IsNullOrWhiteSpace(original))
            {
                continue;
            }

            var translated = await TranslateBlockAsync(
                original,
                "Word 段落",
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);
            bilingualSegments.Add(new BilingualSegment("Word 段落", original, translated));

            var segments = TextDistributionHelper.Distribute(translated, texts.Select(x => Math.Max(1, x.Text.Length)).ToList());
            for (var i = 0; i < texts.Count; i++)
            {
                texts[i].Space = SpaceProcessingModeValues.Preserve;
                texts[i].Text = segments[i];
            }

            var progress = (int)Math.Round((index + 1) * 100d / Math.Max(1, paragraphs.Count));
            await context.ReportProgressAsync(progress, $"Word 段落 {index + 1}/{paragraphs.Count}");
            await context.SaveCheckpointAsync(index + 1, 0, $"Word 段落 {index + 1}/{paragraphs.Count}");
            document.MainDocumentPart?.Document?.Save();
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

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial) =>
        Task.Run(() => progressService.Publish(Path.GetFileName(sourcePath), partial));
}
