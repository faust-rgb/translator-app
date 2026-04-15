using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.IO;
using TranslatorApp.Models;
using UglyToad.PdfPig.Content;

namespace TranslatorApp.Services.Documents;

public sealed class PdfDocumentTranslator(
    ITextTranslationService textTranslationService,
    IAppLogService logService,
    ITranslationProgressService progressService,
    IBilingualExportService bilingualExportService,
    IOcrService ocrService)
    : DocumentTranslatorBase(textTranslationService, logService)
{
    public override bool CanHandle(string extension) => extension == ".pdf";

    public override async Task TranslateAsync(TranslationJobContext context)
    {
        var outputPath = BuildOutputPath(context.Item.SourcePath, context.Settings.Translation.OutputDirectory);
        context.Item.OutputPath = outputPath;

        var startPage = Math.Max(0, context.ResumeUnitIndex);
        var workingPath = outputPath + ".working";
        if (File.Exists(workingPath) && startPage == 0)
        {
            File.Delete(workingPath);
        }

        using var inputPdf = PdfReader.Open(context.Item.SourcePath, PdfDocumentOpenMode.Import);
        using var pig = UglyToad.PdfPig.PdfDocument.Open(context.Item.SourcePath);
        using var outputPdf = new PdfSharp.Pdf.PdfDocument();
        var bilingualSegments = new List<BilingualSegment>();

        if (startPage > 0 && File.Exists(workingPath))
        {
            using var processedPdf = PdfReader.Open(workingPath, PdfDocumentOpenMode.Import);
            for (var importedIndex = 0; importedIndex < Math.Min(startPage, processedPdf.PageCount); importedIndex++)
            {
                outputPdf.AddPage(processedPdf.Pages[importedIndex]);
            }
        }

        for (var pageIndex = startPage; pageIndex < inputPdf.PageCount; pageIndex++)
        {
            await context.PauseController.WaitIfPausedAsync(context.CancellationToken);
            context.CancellationToken.ThrowIfCancellationRequested();

            var importedPage = outputPdf.AddPage(inputPdf.Pages[pageIndex]);
            var pigPage = pig.GetPage(pageIndex + 1);
            var blocks = BuildTextBlocks(pigPage);
            var useOcr = blocks.Count == 0 || pigPage.GetWords().Count() < context.Settings.Ocr.MinimumNativeTextWords;
            if (useOcr)
            {
                var ocrBlocks = await ocrService.RecognizePdfPageAsync(context.Item.SourcePath, pageIndex, context.Settings.Ocr, context.CancellationToken);
                if (ocrBlocks.Count > 0)
                {
                    blocks = ocrBlocks.Select(block => ToPdfTextBlock(block, importedPage.Height.Point)).ToList();
                    Log($"PDF 第 {pageIndex + 1} 页已切换到 OCR 模式。");
                }
            }

            using var graphics = XGraphics.FromPdfPage(importedPage, XGraphicsPdfPageOptions.Append);
            var formatter = new XTextFormatter(graphics);

            for (var blockIndex = 0; blockIndex < blocks.Count; blockIndex++)
            {
                var block = blocks[blockIndex];
                if (string.IsNullOrWhiteSpace(block.Text))
                {
                    continue;
                }

                var translated = await TranslateBlockAsync(
                    block.Text,
                    $"PDF 第 {pageIndex + 1} 页文本块 {blockIndex + 1}",
                    context.Settings,
                    context.PauseController,
                    partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                    context.CancellationToken);

                bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", block.Text, translated));
                DrawTranslatedBlock(graphics, formatter, importedPage.Height.Point, block, translated, context.Settings);
            }

            var progress = (int)Math.Round((pageIndex + 1) * 100d / Math.Max(1, inputPdf.PageCount));
            await context.ReportProgressAsync(progress, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}，文本块 {blocks.Count}");
            outputPdf.Save(workingPath);
            await context.SaveCheckpointAsync(pageIndex + 1, 0, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}");
        }

        outputPdf.Save(outputPath);
        if (File.Exists(workingPath))
        {
            File.Delete(workingPath);
        }

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static List<PdfTextBlock> BuildTextBlocks(Page page)
    {
        var words = page.GetWords()
            .Where(x => !string.IsNullOrWhiteSpace(x.Text))
            .OrderByDescending(x => x.BoundingBox.Top)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        var lines = new List<List<Word>>();
        foreach (var word in words)
        {
            var targetLine = lines.FirstOrDefault(line =>
            {
                var baseline = line.Average(x => x.BoundingBox.Bottom);
                return Math.Abs(baseline - word.BoundingBox.Bottom) < Math.Max(4, GetWordHeight(word) * 0.45);
            });

            if (targetLine is null)
            {
                lines.Add([word]);
            }
            else
            {
                targetLine.Add(word);
            }
        }

        var orderedLines = lines
            .Select(line => line.OrderBy(x => x.BoundingBox.Left).ToList())
            .OrderByDescending(line => line.Max(x => x.BoundingBox.Top))
            .ToList();

        var blocks = new List<PdfTextBlock>();
        foreach (var line in orderedLines)
        {
            var lineText = string.Join(" ", line.Select(x => x.Text));
            if (string.IsNullOrWhiteSpace(lineText))
            {
                continue;
            }

            var rect = GetBoundingRect(line);
            var lineHeight = line.Average(GetWordHeight);
            if (blocks.Count == 0)
            {
                blocks.Add(new PdfTextBlock(lineText, rect, lineHeight, GuessAlignment(page.Width, rect)));
                continue;
            }

            var previous = blocks[^1];
            var verticalGap = previous.Top - rect.Top;
            var similarLeft = Math.Abs(previous.Left - rect.Left) < Math.Max(12, lineHeight);
            var overlap = Math.Min(previous.Right, rect.Right) - Math.Max(previous.Left, rect.Left);
            var hasHorizontalRelation = overlap > 0 || similarLeft;

            if (verticalGap >= -2 && verticalGap < previous.LineHeight * 1.9 && hasHorizontalRelation)
            {
                blocks[^1] = previous.Merge(lineText, rect, lineHeight);
            }
            else
            {
                blocks.Add(new PdfTextBlock(lineText, rect, lineHeight, GuessAlignment(page.Width, rect)));
            }
        }

        return blocks;
    }

    private static void DrawTranslatedBlock(
        XGraphics graphics,
        XTextFormatter formatter,
        double pageHeight,
        PdfTextBlock block,
        string translated,
        Configuration.AppSettings settings)
    {
        var rect = new XRect(
            block.Left - 1,
            pageHeight - block.Top - 1,
            Math.Max(24, block.Width + 2),
            Math.Max(block.Height + 4, block.LineHeight * 1.45));

        var overlay = XColor.FromArgb(248, 255, 255, 255);
        graphics.DrawRectangle(new XSolidBrush(overlay), rect);

        var font = CreateFittedFont(graphics, translated, settings.Translation.OutputFontFamily, Math.Max(block.LineHeight * 0.92, settings.Translation.OutputFontSize), rect.Width, rect.Height);
        formatter.Alignment = block.Alignment;
        formatter.DrawString(translated, font, XBrushes.Black, rect, XStringFormats.TopLeft);
    }

    private static XFont CreateFittedFont(XGraphics graphics, string text, string family, double preferredSize, double maxWidth, double maxHeight)
    {
        var formatter = new XTextFormatter(graphics);
        for (var size = preferredSize; size >= 6; size -= 0.5)
        {
            var font = new XFont(family, size);
            var measured = graphics.MeasureString(text, font);
            if (measured.Width <= maxWidth * 1.3 && measured.Height <= maxHeight * 4)
            {
                return font;
            }
        }

        return new XFont(family, 6);
    }

    private static XParagraphAlignment GuessAlignment(double pageWidth, PdfRect rect)
    {
        var leftSpace = rect.Left;
        var rightSpace = pageWidth - rect.Right;
        if (Math.Abs(leftSpace - rightSpace) < pageWidth * 0.05)
        {
            return XParagraphAlignment.Center;
        }

        return rightSpace < leftSpace * 0.35
            ? XParagraphAlignment.Right
            : XParagraphAlignment.Left;
    }

    private static PdfRect GetBoundingRect(IReadOnlyList<Word> words) =>
        new(
            words.Min(x => x.BoundingBox.Left),
            words.Max(x => x.BoundingBox.Right),
            words.Max(x => x.BoundingBox.Top),
            words.Min(x => x.BoundingBox.Bottom));

    private static double GetWordHeight(Word word) => Math.Max(8, word.BoundingBox.Top - word.BoundingBox.Bottom);

    private static PdfTextBlock ToPdfTextBlock(OcrTextBlock block, double pageHeight) =>
        new(
            block.Text,
            new PdfRect(
                block.Left,
                block.Left + block.Width,
                pageHeight - block.Top,
                pageHeight - (block.Top + block.Height)),
            Math.Max(10, block.Height),
            XParagraphAlignment.Left);

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial) =>
        Task.Run(() => progressService.Publish(Path.GetFileName(sourcePath), partial));

    private sealed record PdfRect(double Left, double Right, double Top, double Bottom)
    {
        public double Width => Right - Left;
        public double Height => Top - Bottom;
    }

    private sealed record PdfTextBlock(string Text, PdfRect Rect, double LineHeight, XParagraphAlignment Alignment)
    {
        public double Left => Rect.Left;
        public double Right => Rect.Right;
        public double Top => Rect.Top;
        public double Bottom => Rect.Bottom;
        public double Width => Rect.Width;
        public double Height => Rect.Height;

        public PdfTextBlock Merge(string lineText, PdfRect rect, double lineHeight) =>
            new(
                $"{Text}\n{lineText}",
                new PdfRect(
                    Math.Min(Rect.Left, rect.Left),
                    Math.Max(Rect.Right, rect.Right),
                    Math.Max(Rect.Top, rect.Top),
                    Math.Min(Rect.Bottom, rect.Bottom)),
                Math.Max(LineHeight, lineHeight),
                Alignment);
    }
}
