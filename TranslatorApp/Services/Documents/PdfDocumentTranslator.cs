using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.IO;
using TranslatorApp.Infrastructure;
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

        var startPage = 0;
        var workingPath = outputPath + ".working";
        if (context.ResumeUnitIndex > 0)
        {
            Log($"PDF 暂不支持从第 {context.ResumeUnitIndex + 1} 页继续写入，已自动从头开始重新生成。");
        }

        if (File.Exists(workingPath))
        {
            File.Delete(workingPath);
        }

        using var inputPdf = PdfReader.Open(context.Item.SourcePath, PdfDocumentOpenMode.Import);
        using var pig = UglyToad.PdfPig.PdfDocument.Open(context.Item.SourcePath);
        using var outputPdf = new PdfSharp.Pdf.PdfDocument();
        var bilingualSegments = new List<BilingualSegment>();

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

            for (var blockIndex = 0; blockIndex < blocks.Count; blockIndex++)
            {
                var block = blocks[blockIndex];
                if (string.IsNullOrWhiteSpace(block.Text))
                {
                    continue;
                }

                try
                {
                    if (IsFormulaLikeBlock(block))
                    {
                        bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", block.Text, block.Text));
                        continue;
                    }

                    var translated = await TranslateBlockAsync(
                        block.Text,
                        $"PDF 第 {pageIndex + 1} 页文本块 {blockIndex + 1}",
                        context.Settings,
                        context.PauseController,
                        partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                        context.CancellationToken);

                    translated ??= string.Empty;
                    bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", block.Text, translated));
                    DrawTranslatedBlock(graphics, importedPage.Width.Point, importedPage.Height.Point, block, translated, context.Settings);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"PDF 第 {pageIndex + 1} 页文本块 {blockIndex + 1} 处理失败：{ex.Message}",
                        ex);
                }
            }

            var progress = (int)Math.Round((pageIndex + 1) * 100d / Math.Max(1, inputPdf.PageCount));
            await context.ReportProgressAsync(progress, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}，文本块 {blocks.Count}");
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
            .Where(x => !IsLikelyMarginalNoise(x, page.Width))
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
            var verticalCenterDistance = Math.Abs(previous.CenterY - rect.CenterY);
            var similarLeft = Math.Abs(previous.Left - rect.Left) < Math.Max(12, lineHeight);
            var similarRight = Math.Abs(previous.Right - rect.Right) < Math.Max(18, lineHeight * 1.5);
            var horizontalOverlap = Math.Min(previous.Right, rect.Right) - Math.Max(previous.Left, rect.Left);
            var overlapRatio = horizontalOverlap / Math.Max(1, Math.Min(previous.Width, rect.Width));
            var hasHorizontalRelation = overlapRatio > 0.45 || similarLeft || similarRight;
            var isCloseLine = verticalCenterDistance < Math.Max(previous.LineHeight, lineHeight) * 1.9;

            if (isCloseLine && hasHorizontalRelation)
            {
                blocks[^1] = previous.Merge(lineText, rect, lineHeight);
            }
            else
            {
                blocks.Add(new PdfTextBlock(lineText, rect, lineHeight, GuessAlignment(page.Width, rect)));
            }
        }

        return MergeRelatedBlocks(blocks, page.Width);
    }

    private static void DrawTranslatedBlock(
        XGraphics graphics,
        double pageWidth,
        double pageHeight,
        PdfTextBlock block,
        string translated,
        Configuration.AppSettings settings)
    {
        var x = Math.Max(2, block.Left - 1);
        var y = Math.Max(2, pageHeight - block.Top - 1);
        var width = Math.Min(Math.Max(24, block.Width + 2), Math.Max(24, pageWidth - x - 2));
        var height = Math.Max(block.Height + 6, block.LineHeight * 1.8);
        var rect = new XRect(x, y, width, height);

        var overlay = XColor.FromArgb(255, 255, 255, 255);
        graphics.DrawRectangle(new XSolidBrush(overlay), rect);

        translated ??= string.Empty;
        var preferredSize = Math.Max(block.LineHeight * 0.92, settings.Translation.OutputFontSize);
        var (font, lines) = FitWrappedText(graphics, translated, settings.Translation.OutputFontFamily, preferredSize, rect.Width, rect.Height);
        var effectiveAlignment = ResolveRenderedAlignment(block, lines, pageWidth);
        DrawWrappedLines(graphics, rect, font, lines, effectiveAlignment);
    }

    private static (XFont Font, IReadOnlyList<string> Lines) FitWrappedText(XGraphics graphics, string text, string family, double preferredSize, double maxWidth, double maxHeight)
    {
        text ??= string.Empty;
        family = string.IsNullOrWhiteSpace(family) ? PdfSharpFontResolver.DefaultFontFamily : family;
        for (var size = preferredSize; size >= 6; size -= 0.5)
        {
            var font = new XFont(family, size);
            var lines = WrapText(graphics, text, font, maxWidth);
            var lineHeight = GetLineHeight(font);
            var totalHeight = lines.Count * lineHeight;
            if (totalHeight <= maxHeight)
            {
                return (font, lines);
            }
        }

        var fallbackFont = new XFont(family, 6);
        return (fallbackFont, WrapText(graphics, text, fallbackFont, maxWidth));
    }

    private static IReadOnlyList<string> WrapText(XGraphics graphics, string text, XFont font, double maxWidth)
    {
        var paragraphs = text.Replace("\r\n", "\n").Split('\n');
        var lines = new List<string>();

        foreach (var paragraph in paragraphs)
        {
            var trimmedParagraph = paragraph.Trim();
            if (string.IsNullOrEmpty(trimmedParagraph))
            {
                lines.Add(string.Empty);
                continue;
            }

            var current = string.Empty;
            foreach (var ch in trimmedParagraph)
            {
                var candidate = string.IsNullOrEmpty(current) ? ch.ToString() : current + ch;
                var measured = graphics.MeasureString(candidate, font);
                if (measured.Width <= maxWidth || string.IsNullOrEmpty(current))
                {
                    current = candidate;
                    continue;
                }

                lines.Add(current.TrimEnd());
                current = char.IsWhiteSpace(ch) ? string.Empty : ch.ToString();
            }

            if (!string.IsNullOrWhiteSpace(current))
            {
                lines.Add(current.TrimEnd());
            }
        }

        return lines.Count == 0 ? [string.Empty] : lines;
    }

    private static void DrawWrappedLines(XGraphics graphics, XRect rect, XFont font, IReadOnlyList<string> lines, XParagraphAlignment alignment)
    {
        var lineHeight = GetLineHeight(font);
        var y = rect.Top;

        foreach (var line in lines)
        {
            if (y + lineHeight > rect.Bottom + 0.5)
            {
                break;
            }

            var rendered = line ?? string.Empty;
            var measured = graphics.MeasureString(rendered, font);
            var x = alignment switch
            {
                XParagraphAlignment.Center => rect.Left + Math.Max(0, (rect.Width - measured.Width) / 2),
                XParagraphAlignment.Right => rect.Right - measured.Width,
                _ => rect.Left
            };

            graphics.DrawString(rendered, font, XBrushes.Black, new XPoint(x, y + font.Size), XStringFormats.Default);
            y += lineHeight;
        }
    }

    private static XParagraphAlignment ResolveRenderedAlignment(PdfTextBlock block, IReadOnlyList<string> lines, double pageWidth)
    {
        if (lines.Count <= 1)
        {
            return block.Alignment;
        }

        var isBodyParagraph = block.Width > pageWidth * 0.45;
        return isBodyParagraph ? XParagraphAlignment.Left : block.Alignment;
    }

    private static double GetLineHeight(XFont font) => font.Size * 1.35;

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

    private static List<PdfTextBlock> MergeRelatedBlocks(IReadOnlyList<PdfTextBlock> blocks, double pageWidth)
    {
        if (blocks.Count <= 1)
        {
            return blocks.ToList();
        }

        var merged = new List<PdfTextBlock> { blocks[0] };
        for (var index = 1; index < blocks.Count; index++)
        {
            var current = blocks[index];
            var previous = merged[^1];

            if (ShouldMergeBlocks(previous, current, pageWidth))
            {
                merged[^1] = previous.Merge(current.Text, current.Rect, current.LineHeight);
            }
            else
            {
                merged.Add(current);
            }
        }

        return merged;
    }

    private static bool ShouldMergeBlocks(PdfTextBlock previous, PdfTextBlock current, double pageWidth)
    {
        if (IsFormulaLikeBlock(previous) || IsFormulaLikeBlock(current))
        {
            return false;
        }

        var lineHeight = Math.Max(previous.LineHeight, current.LineHeight);
        var verticalGap = previous.Bottom - current.Top;
        var alignedLeft = Math.Abs(previous.Left - current.Left) < Math.Max(14, lineHeight * 1.2);
        var alignedRight = Math.Abs(previous.Right - current.Right) < Math.Max(20, lineHeight * 1.8);
        var overlap = Math.Min(previous.Right, current.Right) - Math.Max(previous.Left, current.Left);
        var overlapRatio = overlap / Math.Max(1, Math.Min(previous.Width, current.Width));
        var sameColumn = overlapRatio > 0.55 || alignedLeft || alignedRight;
        var sameAlignment = previous.Alignment == current.Alignment;
        var gapLooksContinuous = verticalGap >= -2 && verticalGap < lineHeight * 1.5;
        var continuationEvidence = EndsWithContinuationCue(previous.Text) || StartsWithContinuationCue(current.Text);

        if (gapLooksContinuous && continuationEvidence)
        {
            var likelySameParagraph = overlapRatio > 0.2 ||
                                      alignedLeft ||
                                      alignedRight ||
                                      current.Left < previous.Left + previous.Width * 0.35;
            if (likelySameParagraph)
            {
                return true;
            }
        }

        if (!(sameColumn && sameAlignment && gapLooksContinuous))
        {
            return false;
        }

        if (LooksLikeCenteredTitleContinuation(previous, current, pageWidth))
        {
            return true;
        }

        if (continuationEvidence)
        {
            return true;
        }

        var paragraphLike = previous.Width > pageWidth * 0.45 && current.Width > pageWidth * 0.45;
        return paragraphLike;
    }

    private static bool LooksLikeCenteredTitleContinuation(PdfTextBlock previous, PdfTextBlock current, double pageWidth) =>
        previous.Alignment == XParagraphAlignment.Center &&
        current.Alignment == XParagraphAlignment.Center &&
        previous.Width < pageWidth * 0.85 &&
        current.Width < pageWidth * 0.85;

    private static bool EndsWithContinuationCue(string text)
    {
        var trimmed = text.TrimEnd();
        if (string.IsNullOrEmpty(trimmed))
        {
            return false;
        }

        var last = trimmed[^1];
        return !".!?。！？:：;；)]}\"'".Contains(last);
    }

    private static bool StartsWithContinuationCue(string text)
    {
        var trimmed = text.TrimStart();
        if (string.IsNullOrEmpty(trimmed))
        {
            return false;
        }

        var first = trimmed[0];
        return char.IsLower(first) ||
               char.IsDigit(first) ||
               ")]}\"',.;:!?".Contains(first) ||
               trimmed.StartsWith("and ", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("or ", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("but ", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("that ", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("which ", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("who ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFormulaLikeBlock(PdfTextBlock block)
    {
        var text = block.Text.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var compact = text.Replace(" ", string.Empty).Replace("\n", string.Empty);
        var hasEquationOperators = compact.Any(ch => "=+-*/^_∑∫∇√≈≤≥∈∀∂".Contains(ch));
        var hasManyMathMarkers = compact.Count(ch => "()[]{}=+-*/^_∑∫∇√≈≤≥∈∀∂".Contains(ch)) >= 2;
        var hasVariablePattern = System.Text.RegularExpressions.Regex.IsMatch(text, @"\b[a-zA-Z]\s*=\s*") ||
                                 System.Text.RegularExpressions.Regex.IsMatch(text, @"arg\s+min", System.Text.RegularExpressions.RegexOptions.IgnoreCase) ||
                                 System.Text.RegularExpressions.Regex.IsMatch(text, @"\bf\s*[∈∉]\s*[A-Za-z]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var mostlyShortSymbols = compact.Length <= 48 && compact.Count(ch => char.IsLetterOrDigit(ch)) < compact.Length * 0.7;

        return (hasEquationOperators && hasManyMathMarkers) || hasVariablePattern || mostlyShortSymbols;
    }

    private static bool IsLikelyMarginalNoise(Word word, double pageWidth)
    {
        var boxWidth = Math.Max(1, word.BoundingBox.Right - word.BoundingBox.Left);
        var boxHeight = Math.Max(1, word.BoundingBox.Top - word.BoundingBox.Bottom);
        var centerX = (word.BoundingBox.Left + word.BoundingBox.Right) / 2;
        var inSideMargin = centerX < pageWidth * 0.1 || centerX > pageWidth * 0.9;
        var looksVertical = boxHeight > boxWidth * 1.4;
        var shortToken = word.Text.Trim().Length <= 3;
        var arxivLike = word.Text.Contains("arXiv", StringComparison.OrdinalIgnoreCase) ||
                        word.Text.Contains("[cs.", StringComparison.OrdinalIgnoreCase);

        return inSideMargin && (looksVertical || shortToken || arxivLike);
    }

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
        public double CenterY => (Top + Bottom) / 2;
    }

    private sealed record PdfTextBlock(string Text, PdfRect Rect, double LineHeight, XParagraphAlignment Alignment)
    {
        public double Left => Rect.Left;
        public double Right => Rect.Right;
        public double Top => Rect.Top;
        public double Bottom => Rect.Bottom;
        public double Width => Rect.Width;
        public double Height => Rect.Height;
        public double CenterY => Rect.CenterY;

        public PdfTextBlock Merge(string lineText, PdfRect rect, double lineHeight) =>
            new(
                MergeText(Text, lineText),
                new PdfRect(
                    Math.Min(Rect.Left, rect.Left),
                    Math.Max(Rect.Right, rect.Right),
                    Math.Max(Rect.Top, rect.Top),
                    Math.Min(Rect.Bottom, rect.Bottom)),
                Math.Max(LineHeight, lineHeight),
                Alignment);

        private static string MergeText(string currentText, string nextLine)
        {
            currentText = currentText.TrimEnd();
            nextLine = nextLine.TrimStart();

            if (string.IsNullOrWhiteSpace(currentText))
            {
                return nextLine;
            }

            if (string.IsNullOrWhiteSpace(nextLine))
            {
                return currentText;
            }

            return EndsWithHyphen(currentText)
                ? string.Concat(currentText.AsSpan(0, currentText.Length - 1), nextLine)
                : $"{currentText} {nextLine}";
        }

        private static bool EndsWithHyphen(string text) =>
            text.EndsWith('-') ||
            text.EndsWith('‐') ||
            text.EndsWith('‑') ||
            text.EndsWith('‒');
    }
}
