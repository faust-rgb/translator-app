using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.IO;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Graphics.Colors;

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

        if (context.ResumeUnitIndex > 0)
        {
            Log($"PDF 暂不支持从第 {context.ResumeUnitIndex + 1} 页继续写入，已自动从头开始重新生成。");
        }

        using var inputPdf = PdfReader.Open(context.Item.SourcePath, PdfDocumentOpenMode.Import);
        using var pig = UglyToad.PdfPig.PdfDocument.Open(context.Item.SourcePath);
        using var outputPdf = new PdfSharp.Pdf.PdfDocument();
        var bilingualSegments = new List<BilingualSegment>();

        for (var pageIndex = 0; pageIndex < inputPdf.PageCount; pageIndex++)
        {
            await context.PauseController.WaitIfPausedAsync(context.CancellationToken);
            context.CancellationToken.ThrowIfCancellationRequested();

            var importedPage = outputPdf.AddPage(inputPdf.Pages[pageIndex]);
            var pigPage = pig.GetPage(pageIndex + 1);
            var blocks = BuildTextBlocks(pigPage);
            var useOcr = ShouldUseOcr(pigPage, blocks, context.Settings.Ocr.MinimumNativeTextWords);
            if (useOcr)
            {
                var ocrBlocks = await ocrService.RecognizePdfPageAsync(context.Item.SourcePath, pageIndex, context.Settings.Ocr, context.CancellationToken);
                if (ocrBlocks.Count > 0)
                {
                    blocks = ocrBlocks
                        .Where(block => ShouldKeepOcrBlock(block, importedPage.Width.Point, importedPage.Height.Point))
                        .Select(block => ToPdfTextBlock(block, importedPage.Width.Point, importedPage.Height.Point))
                        .ToList();
                    Log($"PDF 第 {pageIndex + 1} 页已切换到 OCR 模式。");
                }
            }

            using var graphics = XGraphics.FromPdfPage(importedPage, XGraphicsPdfPageOptions.Append);
            var translatableBlocks = blocks
                .Select((block, index) => new { Block = block, BlockIndex = index })
                .Where(x => !string.IsNullOrWhiteSpace(x.Block.Text))
                .ToList();

            var translationMap = new Dictionary<int, string>();
            var nonFormulaBlocks = translatableBlocks
                .Where(x => !IsFormulaLikeBlock(x.Block))
                .ToList();
            var batchSize = GetBlockTranslationConcurrency(context.Settings);

            for (var batchStart = 0; batchStart < nonFormulaBlocks.Count; batchStart += batchSize)
            {
                var batch = nonFormulaBlocks.Skip(batchStart).Take(batchSize).ToList();
                var translatedBatch = await TranslateBatchAsync(
                    batch.Select(x => new TranslationBlock(x.Block.Text, $"PDF 第 {pageIndex + 1} 页文本块 {x.BlockIndex + 1}")).ToList(),
                    context.Settings,
                    context.PauseController,
                    partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                    context.CancellationToken);

                for (var batchIndex = 0; batchIndex < batch.Count; batchIndex++)
                {
                    translationMap[batch[batchIndex].BlockIndex] = translatedBatch[batchIndex] ?? string.Empty;
                }
            }

            foreach (var entry in translatableBlocks)
            {
                try
                {
                    if (IsFormulaLikeBlock(entry.Block))
                    {
                        bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", entry.Block.Text, entry.Block.Text));
                        continue;
                    }

                    var translated = translationMap.GetValueOrDefault(entry.BlockIndex, string.Empty);
                    bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", entry.Block.Text, translated));
                    DrawTranslatedBlock(graphics, importedPage.Width.Point, importedPage.Height.Point, entry.Block, translated, context.Settings);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"PDF 第 {pageIndex + 1} 页文本块 {entry.BlockIndex + 1} 处理失败：{ex.Message}",
                        ex);
                }
            }

            var progress = (int)Math.Round((pageIndex + 1) * 100d / Math.Max(1, inputPdf.PageCount));
            await context.ReportProgressAsync(progress, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}，文本块 {blocks.Count}");
            await context.SaveCheckpointAsync(pageIndex + 1, 0, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}");
        }

        outputPdf.Save(outputPath);

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static bool ShouldUseOcr(Page page, IReadOnlyList<PdfTextBlock> blocks, int minimumNativeTextWords)
    {
        if (blocks.Count == 0)
        {
            return true;
        }

        var nativeWordCount = page.GetWords().Count();
        if (nativeWordCount < minimumNativeTextWords)
        {
            return true;
        }

        var pageArea = Math.Max(1, page.Width * page.Height);
        var textCoverage = blocks.Sum(block => Math.Max(1, block.Width) * Math.Max(1, block.Height)) / pageArea;
        var sparseNativeText = nativeWordCount < minimumNativeTextWords * 2 &&
                               textCoverage < 0.025 &&
                               blocks.Count <= 3;
        return sparseNativeText;
    }

    // ========================================================================
    // 文本块检测和边界框计算（优化）
    // ========================================================================

    private static List<PdfTextBlock> BuildTextBlocks(Page page)
    {
        var words = page.GetWords()
            .Where(x => !string.IsNullOrWhiteSpace(x.Text))
            .Where(x => !IsLikelyMarginalNoise(x, page.Width))
            .OrderByDescending(x => x.BoundingBox.Top)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        // 第一步：检测多列布局
        var columns = DetectColumns(words, page.Width);

        // 第二步：在每列内部构建行和段落块
        var allBlocks = new List<PdfTextBlock>();
        foreach (var columnWords in columns)
        {
            var columnBlocks = BuildColumnBlocks(columnWords, page.Width);
            allBlocks.AddRange(columnBlocks);
        }

        // 第三步：按阅读顺序排列块（从上到下，从左到右）
        allBlocks.Sort((a, b) =>
        {
            var topDiff = b.Top - a.Top;
            if (Math.Abs(topDiff) > Math.Max(a.LineHeight, b.LineHeight) * 0.5)
            {
                return topDiff > 0 ? 1 : -1;
            }

            return a.Left.CompareTo(b.Left);
        });

        // 第四步：合并相关块（但保护多列结构）
        var mergedBlocks = MergeRelatedBlocks(allBlocks, page.Width);
        return RefineBlockRoles(mergedBlocks, page.Width, page.Height);
    }

    /// <summary>
    /// 检测页面中的多列布局。
    /// 使用单词的水平位置聚类来识别列边界。
    /// </summary>
    private static List<List<Word>> DetectColumns(IReadOnlyList<Word> words, double pageWidth)
    {
        if (words.Count < 5)
        {
            return [words.ToList()];
        }

        // 收集所有单词的左边距，构建直方图来检测列
        var leftEdges = words
            .Select(w => w.BoundingBox.Left)
            .OrderBy(x => x)
            .ToList();

        // 使用间隙检测法：找到左边距的大间隙作为列分隔
        var gaps = new List<(double Position, double GapWidth)>();
        var minColumnGap = pageWidth * 0.04; // 列间距至少为页面宽度的4%

        for (var i = 1; i < leftEdges.Count; i++)
        {
            var gap = leftEdges[i] - leftEdges[i - 1];
            if (gap > minColumnGap)
            {
                gaps.Add((leftEdges[i - 1] + gap / 2, gap));
            }
        }

        // 验证：对每个候选列分隔线，检查是否贯穿页面大部分高度
        var pageTop = words.Max(w => w.BoundingBox.Top);
        var pageBottom = words.Min(w => w.BoundingBox.Bottom);
        var pageContentHeight = Math.Max(1, pageTop - pageBottom);

        var columnBoundaries = new List<double> { 0 };

        foreach (var (position, gapWidth) in gaps.OrderByDescending(g => g.GapWidth).Take(3))
        {
            // 检查分隔线两侧是否都有足够的文本
            var leftWords = words.Where(w => w.BoundingBox.Right < position).ToList();
            var rightWords = words.Where(w => w.BoundingBox.Left > position).ToList();

            if (leftWords.Count < 3 || rightWords.Count < 3)
            {
                continue;
            }

            // 检查两侧文本是否都有足够的垂直跨度
            var leftVertSpan = leftWords.Max(w => w.BoundingBox.Top) - leftWords.Min(w => w.BoundingBox.Bottom);
            var rightVertSpan = rightWords.Max(w => w.BoundingBox.Top) - rightWords.Min(w => w.BoundingBox.Bottom);

            if (leftVertSpan > pageContentHeight * 0.25 && rightVertSpan > pageContentHeight * 0.25)
            {
                columnBoundaries.Add(position);
            }
        }

        columnBoundaries.Add(pageWidth);
        columnBoundaries.Sort();

        if (columnBoundaries.Count <= 2)
        {
            return [words.ToList()];
        }

        // 将单词分配到各列
        var columns = new List<List<Word>>();
        for (var i = 0; i < columnBoundaries.Count - 1; i++)
        {
            var left = columnBoundaries[i];
            var right = columnBoundaries[i + 1];
            var columnCenter = (left + right) / 2;

            var columnWords = words
                .Where(w =>
                {
                    var wordCenter = (w.BoundingBox.Left + w.BoundingBox.Right) / 2;
                    return wordCenter >= left && wordCenter < right;
                })
                .ToList();

            if (columnWords.Count > 0)
            {
                columns.Add(columnWords);
            }
        }

        return columns.Count > 0 ? columns : [words.ToList()];
    }

    /// <summary>
    /// 在单列内部构建文本行和段落块。
    /// 使用基于字形实际尺寸的精确边界框计算。
    /// </summary>
    private static List<PdfTextBlock> BuildColumnBlocks(List<Word> words, double pageWidth)
    {
        words = words
            .OrderByDescending(x => x.BoundingBox.Top)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        var lines = new List<List<Word>>();
        foreach (var word in words)
        {
            var wordHeight = GetWordHeight(word);
            var targetLine = lines.FirstOrDefault(line =>
            {
                // 使用基线对齐来分组行：比较 Bottom（基线近似值）
                var lineBaseline = line.Average(x => x.BoundingBox.Bottom);
                var baselineThreshold = Math.Max(3, wordHeight * 0.4);
                return Math.Abs(lineBaseline - word.BoundingBox.Bottom) < baselineThreshold;
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
            var lineText = BuildLineText(line);
            if (string.IsNullOrWhiteSpace(lineText))
            {
                continue;
            }

            var rect = GetPreciseBoundingRect(line);
            var lineHeight = ComputePreciseLineHeight(line);
            var style = GetTextStyle(line);
            var blockType = DetectBlockType(line, lineText);
            var alignment = GuessAlignment(pageWidth, rect, line);

            if (blocks.Count == 0)
            {
                blocks.Add(new PdfTextBlock(lineText, rect, lineHeight, alignment, style, blockType));
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

            // 不合并不同类型的块（例如列表项和段落）
            var sameType = previous.BlockType == blockType ||
                           previous.BlockType == PdfBlockType.Normal ||
                           blockType == PdfBlockType.Normal;

            if (isCloseLine && hasHorizontalRelation && sameType)
            {
                blocks[^1] = previous.Merge(lineText, rect, lineHeight, style);
            }
            else
            {
                blocks.Add(new PdfTextBlock(lineText, rect, lineHeight, alignment, style, blockType));
            }
        }

        return blocks;
    }

    /// <summary>
    /// 构建行文本，正确处理CJK和拉丁文字间的空格。
    /// </summary>
    private static string BuildLineText(IReadOnlyList<Word> words)
    {
        if (words.Count == 0)
        {
            return string.Empty;
        }

        var result = words[0].Text;
        for (var i = 1; i < words.Count; i++)
        {
            var prevWord = words[i - 1];
            var currWord = words[i];
            var gap = currWord.BoundingBox.Left - prevWord.BoundingBox.Right;
            var avgCharWidth = (GetWordWidth(prevWord) + GetWordWidth(currWord)) /
                               Math.Max(1, prevWord.Text.Length + currWord.Text.Length);

            // 如果两个相邻词之间的间距大于平均字符宽度的一半，则加空格
            // 对于CJK文字，间距通常较小，不需要额外空格
            var prevEndsWithCjk = prevWord.Text.Length > 0 && IsCjkCharacter(prevWord.Text[^1]);
            var currStartsWithCjk = currWord.Text.Length > 0 && IsCjkCharacter(currWord.Text[0]);

            if (prevEndsWithCjk && currStartsWithCjk)
            {
                // CJK之间不加空格
                result += currWord.Text;
            }
            else if (gap > avgCharWidth * 2.5)
            {
                // 大间距可能是特殊分隔
                result += "  " + currWord.Text;
            }
            else
            {
                result += " " + currWord.Text;
            }
        }

        return result;
    }

    /// <summary>
    /// 计算精确的边界框，考虑字体的 ascent/descent。
    /// 使用字母级别的边界信息获取更准确的区域。
    /// </summary>
    private static PdfRect GetPreciseBoundingRect(IReadOnlyList<Word> words)
    {
        // 先获取基于词的边界框
        var left = double.MaxValue;
        var right = double.MinValue;
        var top = double.MinValue;
        var bottom = double.MaxValue;

        foreach (var word in words)
        {
            // 尝试使用字母级别的边界信息获取更精确的区域
            var letters = word.Letters;
            if (letters.Count > 0)
            {
                foreach (var letter in letters)
                {
                    left = Math.Min(left, letter.GlyphRectangle.Left);
                    right = Math.Max(right, letter.GlyphRectangle.Right);
                    top = Math.Max(top, letter.GlyphRectangle.Top);
                    bottom = Math.Min(bottom, letter.GlyphRectangle.Bottom);
                }
            }
            else
            {
                left = Math.Min(left, word.BoundingBox.Left);
                right = Math.Max(right, word.BoundingBox.Right);
                top = Math.Max(top, word.BoundingBox.Top);
                bottom = Math.Min(bottom, word.BoundingBox.Bottom);
            }
        }

        return new PdfRect(left, right, top, bottom);
    }

    /// <summary>
    /// 使用字体度量信息计算更精确的行高。
    /// </summary>
    private static double ComputePreciseLineHeight(IReadOnlyList<Word> words)
    {
        var letters = words.SelectMany(w => w.Letters).ToList();
        if (letters.Count == 0)
        {
            return Math.Max(8, words.Average(GetWordHeight));
        }

        // 使用字母的实际点大小和字形高度来计算行高
        var pointSizes = letters
            .Select(l => l.PointSize > 0 ? l.PointSize : l.FontSize)
            .Where(s => s > 0)
            .ToList();

        if (pointSizes.Count == 0)
        {
            return Math.Max(8, words.Average(GetWordHeight));
        }

        var avgPointSize = pointSizes.Average();

        // 实际字形高度（从基线到顶部和底部）
        var glyphHeights = letters
            .Select(l => l.GlyphRectangle.Top - l.GlyphRectangle.Bottom)
            .Where(h => h > 0)
            .ToList();

        if (glyphHeights.Count > 0)
        {
            var avgGlyphHeight = glyphHeights.Average();
            // 行高 = 最大值（字形高度, 点大小 * 1.2），确保行间不会太紧
            return Math.Max(avgGlyphHeight, avgPointSize * 1.15);
        }

        return Math.Max(8, avgPointSize * 1.2);
    }

    // ========================================================================
    // 覆盖层和文本绘制（优化）
    // ========================================================================

    private static void DrawTranslatedBlock(
        XGraphics graphics,
        double pageWidth,
        double pageHeight,
        PdfTextBlock block,
        string translated,
        Configuration.AppSettings settings)
    {
        translated ??= string.Empty;
        var preferredFamily = ResolvePreferredFontFamily(block.FontFamily, settings.Translation.OutputFontFamily);
        var renderPlan = BuildRenderPlan(graphics, pageWidth, pageHeight, block, translated, preferredFamily, settings);
        graphics.DrawRectangle(new XSolidBrush(renderPlan.OverlayColor), renderPlan.Rect);

        var effectiveAlignment = ResolveRenderedAlignment(block, renderPlan.Lines, pageWidth);
        DrawWrappedLines(graphics, renderPlan.Rect, renderPlan.Font, renderPlan.Lines, effectiveAlignment, block.TextColor, renderPlan.LineHeightRatio, block);
    }

    private static PdfRenderPlan BuildRenderPlan(
        XGraphics graphics,
        double pageWidth,
        double pageHeight,
        PdfTextBlock block,
        string translated,
        string preferredFamily,
        Configuration.AppSettings settings)
    {
        var renderProfile = GetRenderProfile(block, settings);
        var preferredSize = block.FontPointSize > 0
            ? block.FontPointSize
            : Math.Max(block.LineHeight * 0.85, settings.Translation.OutputFontSize);

        foreach (var rect in BuildCandidateRects(block, pageWidth, pageHeight, translated, preferredSize, renderProfile))
        {
            var textAreaWidth = Math.Max(preferredSize * 1.4, rect.Width - renderProfile.InnerHorizontalPadding);
            var textAreaHeight = Math.Max(preferredSize * 1.2, rect.Height - renderProfile.InnerVerticalPadding);
            var layout = TryFitWrappedText(
                graphics,
                translated,
                preferredFamily,
                preferredSize,
                block.Style.IsBold,
                block.Style.IsItalic,
                renderProfile.LineHeightRatio,
                textAreaWidth,
                textAreaHeight,
                renderProfile.MinFontSize);

            if (layout.Fits)
            {
                return new PdfRenderPlan(rect, layout.Font, layout.Lines, layout.LineHeightRatio, EstimateOverlayColor(block, renderProfile, false), false);
            }
        }

        var emergencyRect = BuildCandidateRects(block, pageWidth, pageHeight, translated, preferredSize, renderProfile).Last();
        var emergencyLayout = TryFitWrappedText(
            graphics,
            translated,
            preferredFamily,
            preferredSize,
            block.Style.IsBold,
            block.Style.IsItalic,
            renderProfile.LineHeightRatio * 0.97,
            Math.Max(preferredSize * 1.2, emergencyRect.Width - renderProfile.InnerHorizontalPadding),
            Math.Max(preferredSize, emergencyRect.Height - renderProfile.InnerVerticalPadding),
            Math.Max(4.5, renderProfile.MinFontSize - 1));

        return new PdfRenderPlan(
            emergencyRect,
            emergencyLayout.Font,
            emergencyLayout.Lines,
            emergencyLayout.LineHeightRatio,
            EstimateOverlayColor(block, renderProfile, true),
            true);
    }

    private static IReadOnlyList<XRect> BuildCandidateRects(
        PdfTextBlock block,
        double pageWidth,
        double pageHeight,
        string translated,
        double preferredSize,
        PdfRenderProfile renderProfile)
    {
        var originalLength = Math.Max(1, block.Text.Trim().Length);
        var translatedLength = Math.Max(1, translated.Trim().Length);
        var expansionRatio = translatedLength / (double)originalLength;
        var baseRect = CreateRenderRect(block, pageWidth, pageHeight, preferredSize, renderProfile, 1.0, 1.0);
        var candidates = new List<XRect> { baseRect };

        if (expansionRatio > 1.15 || block.Source == PdfTextSource.Ocr)
        {
            candidates.Add(CreateRenderRect(block, pageWidth, pageHeight, preferredSize, renderProfile, 1.12, 1.18));
        }

        if (expansionRatio > 1.45 || renderProfile.AllowAggressiveExpansion)
        {
            candidates.Add(CreateRenderRect(block, pageWidth, pageHeight, preferredSize, renderProfile, 1.22, 1.35));
        }

        if (block.BlockType is PdfBlockType.Title or PdfBlockType.Caption or PdfBlockType.HeaderFooter)
        {
            candidates.Add(CreateRenderRect(block, pageWidth, pageHeight, preferredSize, renderProfile, 1.3, 1.45));
        }

        return candidates
            .DistinctBy(rect => $"{Math.Round(rect.Left, 2)}:{Math.Round(rect.Top, 2)}:{Math.Round(rect.Width, 2)}:{Math.Round(rect.Height, 2)}")
            .ToList();
    }

    private static XRect CreateRenderRect(
        PdfTextBlock block,
        double pageWidth,
        double pageHeight,
        double preferredSize,
        PdfRenderProfile renderProfile,
        double horizontalScale,
        double verticalScale)
    {
        var fontSize = Math.Max(6, preferredSize);
        var originalLineHeight = Math.Max(8, block.LineHeight);
        var marginLeft = Math.Min(renderProfile.SideMargin * horizontalScale, Math.Max(0, block.Left) * 0.5);
        var marginRight = Math.Min(renderProfile.SideMargin * horizontalScale, Math.Max(0, pageWidth - block.Right) * 0.5);
        var marginTop = Math.Min(renderProfile.TopMargin * verticalScale, Math.Max(originalLineHeight * 0.8, block.Top));
        var marginBottom = Math.Min(renderProfile.BottomMargin * verticalScale, Math.Max(originalLineHeight * 0.8, pageHeight - (pageHeight - block.Bottom)));

        var x = Math.Max(0, block.Left - marginLeft);
        var y = Math.Max(0, pageHeight - block.Top - marginTop);
        var width = Math.Min(block.Width + marginLeft + marginRight, pageWidth - x);
        var height = Math.Min(block.Height + marginTop + marginBottom, pageHeight - y);

        width = Math.Max(width, fontSize * renderProfile.MinWidthInEm);
        height = Math.Max(height, fontSize * renderProfile.MinHeightInEm);
        return new XRect(x, y, width, height);
    }

    private static PdfTextLayout TryFitWrappedText(
        XGraphics graphics,
        string text,
        string family,
        double preferredSize,
        bool isBold,
        bool isItalic,
        double lineHeightRatio,
        double maxWidth,
        double maxHeight,
        double minFontSize)
    {
        text ??= string.Empty;
        family = string.IsNullOrWhiteSpace(family) ? PdfSharpFontResolver.DefaultFontFamily : family;
        var clampedLineHeightRatio = Math.Clamp(lineHeightRatio, 0.95, 2.2);

        foreach (var candidateFamily in GetFontFallbackChain(family))
        {
            for (var size = preferredSize; size >= minFontSize; size -= 0.5)
            {
                try
                {
                    var font = CreateFont(candidateFamily, size, isBold, isItalic);
                    var lines = WrapText(graphics, text, font, maxWidth);
                    var lineHeight = GetLineHeight(font, clampedLineHeightRatio);
                    var totalHeight = lines.Count * lineHeight;
                    if (totalHeight <= maxHeight)
                    {
                        return new PdfTextLayout(font, lines, clampedLineHeightRatio, true);
                    }
                }
                catch
                {
                    // 某些 PDF 提取出的字体名在本机不可用，继续尝试下一种候选字体。
                }
            }
        }

        var fallbackFont = CreateFont(PdfSharpFontResolver.DefaultFontFamily, Math.Max(4.5, minFontSize), isBold, isItalic);
        return new PdfTextLayout(fallbackFont, WrapText(graphics, text, fallbackFont, maxWidth), clampedLineHeightRatio, false);
    }

    private static IEnumerable<string> GetFontFallbackChain(string family)
    {
        yield return family;

        if (!string.Equals(family, PdfSharpFontResolver.DefaultFontFamily, StringComparison.OrdinalIgnoreCase))
        {
            yield return PdfSharpFontResolver.DefaultFontFamily;
        }
    }

    private static XFont CreateFont(string family, double size, bool isBold, bool isItalic)
    {
        var style = XFontStyleEx.Regular;

        if (isBold && isItalic)
        {
            style = XFontStyleEx.BoldItalic;
        }
        else if (isBold)
        {
            style = XFontStyleEx.Bold;
        }
        else if (isItalic)
        {
            style = XFontStyleEx.Italic;
        }

        return new XFont(family, size, style);
    }

    // ========================================================================
    // 文本换行（优化：批量测量 + CJK换行规则）
    // ========================================================================

    private static IReadOnlyList<string> WrapText(XGraphics graphics, string text, XFont font, double maxWidth)
    {
        var paragraphs = text.Replace("\r\n", "\n").Split('\n');
        var lines = new List<string>();

        // 预估平均字符宽度，用于批量估算优化
        var avgCharWidth = EstimateAverageCharWidth(graphics, font);

        foreach (var paragraph in paragraphs)
        {
            var trimmedParagraph = paragraph.Trim();
            if (string.IsNullOrEmpty(trimmedParagraph))
            {
                lines.Add(string.Empty);
                continue;
            }

            // 快速路径：如果整段文本宽度不超过限制，直接添加
            var fullWidth = graphics.MeasureString(trimmedParagraph, font).Width;
            if (fullWidth <= maxWidth)
            {
                lines.Add(trimmedParagraph);
                continue;
            }

            // 使用优化的换行算法
            WrapParagraph(graphics, font, trimmedParagraph, maxWidth, avgCharWidth, lines);
        }

        return lines.Count == 0 ? [string.Empty] : lines;
    }

    /// <summary>
    /// 估算平均字符宽度，用于优化换行计算。
    /// </summary>
    private static double EstimateAverageCharWidth(XGraphics graphics, XFont font)
    {
        // 用一个包含常见字符的样本字符串来估算
        const string sample = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789中文字符测试";
        var width = graphics.MeasureString(sample, font).Width;
        return width / sample.Length;
    }

    /// <summary>
    /// 优化的段落换行算法。
    /// 使用批量估算减少 MeasureString 调用次数，并遵循 CJK 换行规则。
    /// </summary>
    private static void WrapParagraph(XGraphics graphics, XFont font, string text, double maxWidth, double avgCharWidth, List<string> lines)
    {
        var startIndex = 0;

        while (startIndex < text.Length)
        {
            // 估算这一行大约能放多少字符
            var estimatedChars = Math.Max(1, (int)(maxWidth / avgCharWidth));
            var endIndex = Math.Min(startIndex + estimatedChars, text.Length);

            // 二分搜索找到实际能放下的位置
            var low = startIndex + 1;
            var high = endIndex;

            // 先检查估算位置是否能放下
            var testStr = text[startIndex..Math.Min(high, text.Length)];
            var testWidth = graphics.MeasureString(testStr, font).Width;

            if (testWidth <= maxWidth && high >= text.Length)
            {
                // 整段剩余文本都能放下
                lines.Add(text[startIndex..].TrimEnd());
                break;
            }

            if (testWidth <= maxWidth)
            {
                // 估算偏少，向后扩展
                low = high;
                high = Math.Min(high + estimatedChars / 2, text.Length);
            }
            else
            {
                // 估算偏多，向前收缩
                high = endIndex;
            }

            // 二分搜索精确位置
            var fitEnd = startIndex + 1; // 至少放一个字符
            while (low <= high && low <= text.Length)
            {
                var mid = (low + high) / 2;
                var segment = text[startIndex..mid];
                var segWidth = graphics.MeasureString(segment, font).Width;

                if (segWidth <= maxWidth)
                {
                    fitEnd = mid;
                    low = mid + 1;
                }
                else
                {
                    high = mid - 1;
                }
            }

            // 在合法位置断行（CJK/空格/标点）
            var breakPos = FindBestBreakPosition(text, startIndex, fitEnd);
            var lineEnd = breakPos > startIndex ? breakPos : fitEnd;

            lines.Add(text[startIndex..lineEnd].TrimEnd());
            startIndex = lineEnd;

            // 跳过行首空格（对于西文文本）
            while (startIndex < text.Length && text[startIndex] == ' ')
            {
                startIndex++;
            }
        }
    }

    /// <summary>
    /// 在指定范围内查找最佳换行位置，遵循 CJK 换行规则。
    /// </summary>
    private static int FindBestBreakPosition(string text, int startIndex, int maxEnd)
    {
        if (maxEnd >= text.Length)
        {
            return text.Length;
        }

        // 从 maxEnd 向前搜索最佳断点
        var searchStart = Math.Max(startIndex + 1, maxEnd - Math.Max(8, (maxEnd - startIndex) / 3));

        for (var i = maxEnd; i >= searchStart; i--)
        {
            if (i >= text.Length)
            {
                continue;
            }

            var ch = text[i];
            var prevCh = i > 0 ? text[i - 1] : '\0';

            // 1. 空格后断行（西文词边界）
            if (ch == ' ')
            {
                return i;
            }

            // 2. CJK字符前可以断行（除非是行首禁止字符）
            if (IsCjkCharacter(ch) && !IsLineStartProhibited(ch))
            {
                return i;
            }

            // 3. CJK字符后可以断行（除非后面是行尾禁止字符）
            if (IsCjkCharacter(prevCh) && !IsLineEndProhibited(prevCh) && i < text.Length && !IsLineStartProhibited(ch))
            {
                return i;
            }

            // 4. 标点符号后断行（但注意行首禁止标点）
            if (IsBreakableAfterPunctuation(prevCh) && !IsLineStartProhibited(ch))
            {
                return i;
            }
        }

        // 未找到好的断点，强制在 maxEnd 断行
        return maxEnd;
    }

    /// <summary>
    /// 行首禁止字符（不能出现在行首的标点）。
    /// </summary>
    private static bool IsLineStartProhibited(char ch)
    {
        return "!),.:;?]}。，、；：？！）》」』】〉〗〙〛·…—～".Contains(ch);
    }

    /// <summary>
    /// 行尾禁止字符（不能出现在行尾的标点）。
    /// </summary>
    private static bool IsLineEndProhibited(char ch)
    {
        return "([{（《「『【〈〖〘〚".Contains(ch);
    }

    /// <summary>
    /// 可以在该标点后断行。
    /// </summary>
    private static bool IsBreakableAfterPunctuation(char ch)
    {
        return ",，.。;；!！?？:：、)]}>）》」』】〉…—～\"'".Contains(ch);
    }

    private static bool IsCjkCharacter(char ch)
    {
        // CJK统一表意文字基本区、扩展区A、兼容表意文字
        return (ch >= '\u4E00' && ch <= '\u9FFF') ||
               (ch >= '\u3400' && ch <= '\u4DBF') ||
               (ch >= '\uF900' && ch <= '\uFAFF') ||
               // 日文平假名和片假名
               (ch >= '\u3040' && ch <= '\u309F') ||
               (ch >= '\u30A0' && ch <= '\u30FF') ||
               // CJK全角标点
               (ch >= '\uFF00' && ch <= '\uFFEF') ||
               // CJK标点符号
               (ch >= '\u3000' && ch <= '\u303F');
    }

    private static void DrawWrappedLines(
        XGraphics graphics,
        XRect rect,
        XFont font,
        IReadOnlyList<string> lines,
        XParagraphAlignment alignment,
        XColor textColor,
        double lineHeightRatio,
        PdfTextBlock block)
    {
        var lineHeight = GetLineHeight(font, lineHeightRatio);
        var totalTextHeight = lines.Count * lineHeight;

        // 垂直居中对齐（如果文本比区域小，则垂直居中）
        var y = rect.Top;
        if (totalTextHeight < rect.Height - lineHeight * 0.3)
        {
            y = rect.Top + (rect.Height - totalTextHeight) / 2;
        }

        var brush = new XSolidBrush(textColor);
        var innerLeft = rect.Left + font.Size * 0.05; // 微小内边距
        var innerWidth = rect.Width - font.Size * 0.1;
        var hangingIndent = GetHangingIndent(block, font, graphics);

        for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
        {
            var line = lines[lineIndex];
            if (y + lineHeight > rect.Bottom + 0.5)
            {
                break;
            }

            var rendered = line ?? string.Empty;
            var measured = graphics.MeasureString(rendered, font);
            var leftOffset = lineIndex > 0 ? hangingIndent : 0;
            var availableWidth = Math.Max(0, innerWidth - leftOffset);
            var x = alignment switch
            {
                XParagraphAlignment.Center => innerLeft + Math.Max(0, (innerWidth - measured.Width) / 2),
                XParagraphAlignment.Right => rect.Right - measured.Width - font.Size * 0.05,
                _ => innerLeft + leftOffset
            };

            if (alignment == XParagraphAlignment.Left && measured.Width > availableWidth && lineIndex > 0)
            {
                x = innerLeft + Math.Max(0, hangingIndent * 0.5);
            }

            graphics.DrawString(rendered, font, brush, new XPoint(x, y + font.Size), XStringFormats.Default);
            y += lineHeight;
        }
    }

    private static XParagraphAlignment ResolveRenderedAlignment(PdfTextBlock block, IReadOnlyList<string> lines, double pageWidth)
    {
        if (lines.Count <= 1)
        {
            return block.BlockType == PdfBlockType.Title ? XParagraphAlignment.Center : block.Alignment;
        }

        if (block.BlockType == PdfBlockType.Title)
        {
            return XParagraphAlignment.Center;
        }

        if (block.BlockType is PdfBlockType.Caption or PdfBlockType.Footnote or PdfBlockType.ListItem)
        {
            return XParagraphAlignment.Left;
        }

        var isBodyParagraph = block.Width > pageWidth * 0.45;
        return isBodyParagraph ? XParagraphAlignment.Left : block.Alignment;
    }

    private static double GetLineHeight(XFont font, double lineHeightRatio) =>
        font.Size * Math.Clamp(lineHeightRatio, 1.0, 1.8);

    private static double GetHangingIndent(PdfTextBlock block, XFont font, XGraphics graphics)
    {
        if (block.BlockType != PdfBlockType.ListItem)
        {
            return 0;
        }

        var prefix = ExtractListMarker(block.Text);
        if (string.IsNullOrWhiteSpace(prefix))
        {
            return font.Size * 1.2;
        }

        return Math.Min(font.Size * 2.8, graphics.MeasureString(prefix + " ", font).Width);
    }

    private static string ExtractListMarker(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var match = System.Text.RegularExpressions.Regex.Match(text, @"^\s*((\d{1,3}[\.\)、])|([\(（]\d{1,3}[\)）])|([a-zA-Z][\.\)])|([•\-·◦▪■★→►▸‣⦿⁃]))");
        return match.Success ? match.Value.Trim() : string.Empty;
    }

    private static PdfRenderProfile GetRenderProfile(PdfTextBlock block, Configuration.AppSettings settings)
    {
        var baseFontSize = block.FontPointSize > 0
            ? block.FontPointSize
            : Math.Max(block.LineHeight * 0.85, settings.Translation.OutputFontSize);
        var originalLineHeight = Math.Max(8, block.LineHeight);
        var lineHeightRatio = block.FontPointSize > 0
            ? Math.Clamp(originalLineHeight / Math.Max(1, block.FontPointSize), 1.0, 1.95)
            : 1.35;

        var sideMargin = baseFontSize * 0.08;
        var topMargin = baseFontSize * 0.15;
        var bottomMargin = baseFontSize * 0.1;
        var innerHorizontalPadding = baseFontSize * 0.18;
        var innerVerticalPadding = baseFontSize * 0.12;
        var minFontSize = 6d;
        var minWidthInEm = 2d;
        var minHeightInEm = 1.3d;
        var allowAggressiveExpansion = false;

        if (block.Source == PdfTextSource.Ocr)
        {
            sideMargin *= 1.25;
            topMargin *= 1.35;
            bottomMargin *= 1.35;
            lineHeightRatio = Math.Max(lineHeightRatio, 1.28);
            allowAggressiveExpansion = true;
        }

        switch (block.BlockType)
        {
            case PdfBlockType.Title:
                topMargin *= 1.5;
                bottomMargin *= 1.4;
                sideMargin *= 1.2;
                lineHeightRatio = Math.Max(1.18, lineHeightRatio * 0.98);
                minFontSize = 7;
                minHeightInEm = 1.8;
                allowAggressiveExpansion = true;
                break;
            case PdfBlockType.Caption:
                topMargin *= 1.2;
                bottomMargin *= 1.3;
                sideMargin *= 1.1;
                lineHeightRatio = Math.Max(1.2, lineHeightRatio);
                minFontSize = 6;
                minHeightInEm = 1.55;
                break;
            case PdfBlockType.ListItem:
                sideMargin *= 1.15;
                lineHeightRatio = Math.Max(1.2, lineHeightRatio);
                minHeightInEm = 1.5;
                break;
            case PdfBlockType.Code:
                sideMargin *= 1.05;
                lineHeightRatio = Math.Max(1.12, lineHeightRatio);
                innerHorizontalPadding *= 1.2;
                minFontSize = 5.5;
                break;
            case PdfBlockType.TableRow:
                sideMargin *= 1.25;
                topMargin *= 1.15;
                bottomMargin *= 1.15;
                lineHeightRatio = Math.Max(1.15, lineHeightRatio);
                allowAggressiveExpansion = true;
                break;
            case PdfBlockType.HeaderFooter:
                topMargin *= 1.1;
                bottomMargin *= 1.1;
                minFontSize = 5.5;
                minHeightInEm = 1.2;
                break;
            case PdfBlockType.Footnote:
                topMargin *= 1.15;
                bottomMargin *= 1.25;
                lineHeightRatio = Math.Max(1.15, lineHeightRatio);
                minFontSize = 5.5;
                minHeightInEm = 1.35;
                break;
        }

        return new PdfRenderProfile(
            sideMargin,
            topMargin,
            bottomMargin,
            innerHorizontalPadding,
            innerVerticalPadding,
            lineHeightRatio,
            minFontSize,
            minWidthInEm,
            minHeightInEm,
            allowAggressiveExpansion);
    }

    private static XColor EstimateOverlayColor(PdfTextBlock block, PdfRenderProfile profile, bool emergency)
    {
        var brightness = (block.TextColor.R + block.TextColor.G + block.TextColor.B) / (3d * 255);
        if (brightness > 0.8)
        {
            var baseShade = emergency ? 36 : 28;
            return XColor.FromArgb(255, (byte)baseShade, (byte)baseShade, (byte)(baseShade + 6));
        }

        if (block.BlockType == PdfBlockType.TableRow)
        {
            var shade = emergency ? 248 : 244;
            return XColor.FromArgb(255, (byte)shade, (byte)shade, (byte)(shade - 2));
        }

        if (block.BlockType is PdfBlockType.Code or PdfBlockType.Caption)
        {
            var shade = emergency ? 252 : 250;
            return XColor.FromArgb(255, (byte)shade, (byte)shade, (byte)(shade - 2));
        }

        if (profile.AllowAggressiveExpansion && emergency)
        {
            return XColor.FromArgb(255, 252, 252, 252);
        }

        return XColor.FromArgb(255, 255, 255, 255);
    }

    // ========================================================================
    // 特殊元素检测（表格、列表、代码块）
    // ========================================================================

    /// <summary>
    /// 检测文本块类型：普通文本、列表项、代码块、表格行。
    /// </summary>
    private static PdfBlockType DetectBlockType(IReadOnlyList<Word> words, string lineText)
    {
        var trimmed = lineText.Trim();

        // 检测列表项
        if (IsListItem(trimmed))
        {
            return PdfBlockType.ListItem;
        }

        // 检测代码行（使用等宽字体的文本）
        if (IsCodeLine(words, trimmed))
        {
            return PdfBlockType.Code;
        }

        // 检测表格行（包含多个对齐的列）
        if (IsTableRow(words))
        {
            return PdfBlockType.TableRow;
        }

        return PdfBlockType.Normal;
    }

    /// <summary>
    /// 检测是否是列表项。
    /// </summary>
    private static bool IsListItem(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        // 检测有序列表：1. 2. 3. 或 (1) (2)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*(\d{1,3}[\.\)、]|[\(（]\d{1,3}[\)）])\s"))
        {
            return true;
        }

        // 检测无序列表：• - · ◦ ▪ ■ ★ → 等
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*[•\-·◦▪■★→►▸‣⦿⁃]\s"))
        {
            return true;
        }

        // 检测字母列表：a. b. c. 或 A. B. C.
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*[a-zA-Z][\.\)]\s"))
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// 检测是否是代码行。
    /// </summary>
    private static bool IsCodeLine(IReadOnlyList<Word> words, string lineText)
    {
        // 检查字体：代码通常使用等宽字体
        var monoFontCount = words
            .SelectMany(w => w.Letters)
            .Count(l => IsMonospaceFont(l.FontName));
        var totalLetters = words.Sum(w => w.Letters.Count);

        if (totalLetters > 0 && (double)monoFontCount / totalLetters > 0.7)
        {
            return true;
        }

        // 检查内容特征：缩进 + 包含代码标记
        var hasCodeMarkers = lineText.Contains("()") ||
                             lineText.Contains("{}") ||
                             lineText.Contains("[]") ||
                             lineText.Contains("=>") ||
                             lineText.Contains("->") ||
                             lineText.Contains("::") ||
                             lineText.Contains("//") ||
                             lineText.Contains("/*") ||
                             lineText.Contains("#include") ||
                             lineText.Contains("import ") ||
                             lineText.Contains("def ") ||
                             lineText.Contains("class ") ||
                             lineText.Contains("function ");

        var hasSignificantIndent = lineText.Length > 0 && lineText.Length - lineText.TrimStart().Length >= 4;

        return hasCodeMarkers && hasSignificantIndent;
    }

    /// <summary>
    /// 检测是否是等宽字体。
    /// </summary>
    private static bool IsMonospaceFont(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return false;
        }

        var normalized = fontName.ToLowerInvariant();
        return normalized.Contains("courier") ||
               normalized.Contains("consolas") ||
               normalized.Contains("mono") ||
               normalized.Contains("menlo") ||
               normalized.Contains("fira code") ||
               normalized.Contains("source code") ||
               normalized.Contains("jetbrains") ||
               normalized.Contains("inconsolata") ||
               normalized.Contains("hack");
    }

    /// <summary>
    /// 检测是否是表格行（单词间有规律的大间距）。
    /// </summary>
    private static bool IsTableRow(IReadOnlyList<Word> words)
    {
        if (words.Count < 2)
        {
            return false;
        }

        // 计算相邻单词间的间距
        var gaps = new List<double>();
        var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();

        for (var i = 1; i < sortedWords.Count; i++)
        {
            var gap = sortedWords[i].BoundingBox.Left - sortedWords[i - 1].BoundingBox.Right;
            gaps.Add(gap);
        }

        if (gaps.Count < 2)
        {
            return false;
        }

        // 表格行通常有几个明显的大间距（列分隔）
        var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Right - w.BoundingBox.Left);
        var largeGaps = gaps.Count(g => g > avgWordWidth * 1.5);

        return largeGaps >= 2 && largeGaps >= gaps.Count * 0.3;
    }

    private static List<PdfTextBlock> RefineBlockRoles(IReadOnlyList<PdfTextBlock> blocks, double pageWidth, double pageHeight)
    {
        if (blocks.Count == 0)
        {
            return [];
        }

        var medianFontSize = blocks
            .Select(x => x.FontPointSize > 0 ? x.FontPointSize : x.LineHeight)
            .OrderBy(x => x)
            .ElementAt(blocks.Count / 2);

        var refined = new List<PdfTextBlock>(blocks.Count);
        foreach (var block in blocks)
        {
            var blockType = block.BlockType;
            var trimmed = block.Text.Trim();
            var topRatio = pageHeight <= 0 ? 0 : block.Top / pageHeight;
            var bottomRatio = pageHeight <= 0 ? 0 : block.Bottom / pageHeight;
            var fontSize = block.FontPointSize > 0 ? block.FontPointSize : block.LineHeight;
            var shortText = trimmed.Length <= 80;
            var tinyText = trimmed.Length <= 24;
            var centered = block.Alignment == XParagraphAlignment.Center;

            if (IsCaptionText(trimmed))
            {
                blockType = PdfBlockType.Caption;
            }
            else if (IsHeaderFooterText(trimmed, topRatio, bottomRatio, tinyText))
            {
                blockType = PdfBlockType.HeaderFooter;
            }
            else if (IsFootnoteText(trimmed, bottomRatio, fontSize, medianFontSize))
            {
                blockType = PdfBlockType.Footnote;
            }
            else if (IsTitleLikeBlock(block, pageWidth, topRatio, medianFontSize, shortText, centered))
            {
                blockType = PdfBlockType.Title;
            }

            refined.Add(block.WithBlockType(blockType));
        }

        return refined;
    }

    private static bool IsCaptionText(string text) =>
        System.Text.RegularExpressions.Regex.IsMatch(
            text,
            @"^(Figure|Fig\.?|Table|Chart|Source|Appendix|Algorithm|图|表|附录|来源)\s*[\dA-Za-z\.\-:：]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    private static bool IsHeaderFooterText(string text, double topRatio, double bottomRatio, bool tinyText)
    {
        if (!tinyText)
        {
            return false;
        }

        var nearTop = topRatio > 0.9;
        var nearBottom = bottomRatio < 0.1;
        var pageMarker = System.Text.RegularExpressions.Regex.IsMatch(text, @"^(page\s+\d+|\d+\s*/\s*\d+|\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return (nearTop || nearBottom) && (pageMarker || text.Length <= 18);
    }

    private static bool IsFootnoteText(string text, double bottomRatio, double fontSize, double medianFontSize)
    {
        if (bottomRatio >= 0.18 || fontSize > medianFontSize * 0.92)
        {
            return false;
        }

        return System.Text.RegularExpressions.Regex.IsMatch(text, @"^(\d+[\.\)]|\*|†|‡|\[\d+\])\s") ||
               text.Length <= 100;
    }

    private static bool IsTitleLikeBlock(PdfTextBlock block, double pageWidth, double topRatio, double medianFontSize, bool shortText, bool centered)
    {
        var fontSize = block.FontPointSize > 0 ? block.FontPointSize : block.LineHeight;
        var wideEnough = block.Width < pageWidth * 0.8;
        var prominentFont = fontSize >= medianFontSize * 1.15;
        var nearTop = topRatio > 0.58;
        var punctuationLight = !EndsWithSentencePunctuation(block.Text);
        return shortText && wideEnough && punctuationLight && (centered || (nearTop && prominentFont));
    }

    private static bool ShouldKeepOcrBlock(OcrTextBlock block, double pageWidth, double pageHeight)
    {
        if (string.IsNullOrWhiteSpace(block.Text))
        {
            return false;
        }

        var trimmed = block.Text.Trim();
        var centerX = block.Left + block.Width / 2;
        var inMargin = centerX < pageWidth * 0.08 || centerX > pageWidth * 0.92;
        var nearHeaderFooter = block.Top < pageHeight * 0.06 || (block.Top + block.Height) > pageHeight * 0.94;
        var likelyPageNumber = System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^(\d+|page\s+\d+|\d+\s*/\s*\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var tooSmall = trimmed.Length <= 2 && block.Width < pageWidth * 0.08;
        return !(likelyPageNumber || (inMargin && tooSmall) || (nearHeaderFooter && trimmed.Length <= 10));
    }

    // ========================================================================
    // 对齐方式推断
    // ========================================================================

    private static XParagraphAlignment GuessAlignment(double pageWidth, PdfRect rect, IReadOnlyList<Word> lineWords)
    {
        // 计算文本在页面中的相对位置
        var leftSpace = rect.Left;
        var rightSpace = pageWidth - rect.Right;
        var centerOffset = Math.Abs(leftSpace - rightSpace);
        var tolerance = pageWidth * 0.03; // 3% 容差

        // 1. 首先检查是否明显居中（左右边距接近相等）
        if (centerOffset < tolerance)
        {
            return XParagraphAlignment.Center;
        }

        // 2. 检查是否右对齐（右边距很小）
        if (rightSpace < tolerance && leftSpace > pageWidth * 0.15)
        {
            return XParagraphAlignment.Right;
        }

        // 3. 检查是否左对齐（左边距很小或符合常规缩进）
        if (leftSpace < tolerance || (leftSpace < pageWidth * 0.12 && rightSpace > pageWidth * 0.1))
        {
            return XParagraphAlignment.Left;
        }

        // 4. 对于短文本（可能是标题），使用更宽松的居中判断
        var textWidthRatio = rect.Width / pageWidth;
        if (textWidthRatio < 0.6 && centerOffset < pageWidth * 0.08)
        {
            return XParagraphAlignment.Center;
        }

        // 5. 默认左对齐
        return XParagraphAlignment.Left;
    }

    private static PdfRect GetBoundingRect(IReadOnlyList<Word> words) =>
        new(
            words.Min(x => x.BoundingBox.Left),
            words.Max(x => x.BoundingBox.Right),
            words.Max(x => x.BoundingBox.Top),
            words.Min(x => x.BoundingBox.Bottom));

    private static double GetWordHeight(Word word) => Math.Max(8, word.BoundingBox.Top - word.BoundingBox.Bottom);
    private static double GetWordWidth(Word word) => Math.Max(1, word.BoundingBox.Right - word.BoundingBox.Left);

    // ========================================================================
    // 文本块合并
    // ========================================================================

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
                merged[^1] = previous.Merge(current.Text, current.Rect, current.LineHeight, current.Style);
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
        // 公式块不合并
        if (IsFormulaLikeBlock(previous) || IsFormulaLikeBlock(current))
        {
            return false;
        }

        // 不同类型的块不合并（如列表和段落）
        if (previous.BlockType != PdfBlockType.Normal && current.BlockType != PdfBlockType.Normal &&
            previous.BlockType != current.BlockType)
        {
            return false;
        }

        if (previous.BlockType is PdfBlockType.Title or PdfBlockType.Caption or PdfBlockType.HeaderFooter or PdfBlockType.Footnote ||
            current.BlockType is PdfBlockType.Title or PdfBlockType.Caption or PdfBlockType.HeaderFooter or PdfBlockType.Footnote)
        {
            return previous.BlockType == current.BlockType &&
                   Math.Abs(previous.Left - current.Left) < Math.Max(previous.LineHeight, current.LineHeight) &&
                   Math.Abs(previous.Right - current.Right) < Math.Max(previous.LineHeight * 1.5, current.LineHeight * 1.5) &&
                   previous.Bottom - current.Top < Math.Max(previous.LineHeight, current.LineHeight) * 1.4;
        }

        // 代码块之间可以合并
        if (previous.BlockType == PdfBlockType.Code && current.BlockType == PdfBlockType.Code)
        {
            var lineHeight = Math.Max(previous.LineHeight, current.LineHeight);
            var verticalGap = previous.Bottom - current.Top;
            return verticalGap >= -2 && verticalGap < lineHeight * 1.5;
        }

        // 表格行之间可以合并
        if (previous.BlockType == PdfBlockType.TableRow && current.BlockType == PdfBlockType.TableRow)
        {
            var lineHeight = Math.Max(previous.LineHeight, current.LineHeight);
            var verticalGap = previous.Bottom - current.Top;
            return verticalGap >= -2 && verticalGap < lineHeight * 2.0;
        }

        // 字体大小差异过大不合并（可能是标题和正文）
        var fontSizeRatio = Math.Max(previous.FontPointSize, 1) / Math.Max(current.FontPointSize, 1);
        if (fontSizeRatio > 1.5 || fontSizeRatio < 0.67)
        {
            return false;
        }

        var mergeLineHeight = Math.Max(previous.LineHeight, current.LineHeight);
        var mergeVerticalGap = previous.Bottom - current.Top;

        // 水平对齐检查
        var alignedLeft = Math.Abs(previous.Left - current.Left) < Math.Max(14, mergeLineHeight * 1.2);
        var alignedRight = Math.Abs(previous.Right - current.Right) < Math.Max(20, mergeLineHeight * 1.8);
        var overlap = Math.Min(previous.Right, current.Right) - Math.Max(previous.Left, current.Left);
        var overlapRatio = overlap / Math.Max(1, Math.Min(previous.Width, current.Width));
        var sameColumn = overlapRatio > 0.55 || alignedLeft || alignedRight;
        var sameAlignment = previous.Alignment == current.Alignment;

        // 垂直间距检查
        var gapLooksContinuous = mergeVerticalGap >= -2 && mergeVerticalGap < mergeLineHeight * 1.5;

        // 缩进变化检查（可能是新段落开始）
        var indentChange = Math.Abs(previous.Left - current.Left);
        var significantIndentChange = indentChange > mergeLineHeight * 1.5 && previous.Width < pageWidth * 0.7;

        // 如果缩进显著变化，可能是新段落，不合并
        if (significantIndentChange && !alignedLeft)
        {
            return false;
        }

        // 连续性线索
        var continuationEvidence = EndsWithContinuationCue(previous.Text) || StartsWithContinuationCue(current.Text);

        // 情况1：明显连续（有连字符或连接词，且在同一列）
        if (gapLooksContinuous && continuationEvidence && sameColumn)
        {
            return true;
        }

        // 情况2：同一列、相同对齐、间距连续
        if (!(sameColumn && sameAlignment && gapLooksContinuous))
        {
            return false;
        }

        // 情况3：居中标题连续
        if (LooksLikeCenteredTitleContinuation(previous, current, pageWidth))
        {
            return true;
        }

        // 情况4：段落样式（宽度超过页面45%）
        var paragraphLike = previous.Width > pageWidth * 0.45 && current.Width > pageWidth * 0.45;

        // 情况5：检查行高一致性（多行段落通常有相似的行高）
        var consistentLineHeight = Math.Abs(previous.LineHeight - current.LineHeight) < mergeLineHeight * 0.3;

        // 情况6：检查文本内容特征
        var previousEndsWithPunctuation = EndsWithSentencePunctuation(previous.Text);
        var currentStartsWithLowercase = current.Text.TrimStart().Length > 0 && char.IsLower(current.Text.TrimStart()[0]);

        // 如果前一段以句号结束，且当前段以小写开头，很可能是同一段落
        if (previousEndsWithPunctuation && currentStartsWithLowercase && gapLooksContinuous)
        {
            return true;
        }

        // 如果前一段不以句号结束，且在同一列，可能是同一段落
        if (!previousEndsWithPunctuation && sameColumn && gapLooksContinuous && consistentLineHeight)
        {
            return true;
        }

        return paragraphLike && consistentLineHeight;
    }

    private static bool EndsWithSentencePunctuation(string text)
    {
        var trimmed = text.TrimEnd();
        if (string.IsNullOrEmpty(trimmed))
        {
            return false;
        }

        var last = trimmed[^1];
        return ".!?。！？".Contains(last);
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

    // ========================================================================
    // 公式检测
    // ========================================================================

    private static bool IsFormulaLikeBlock(PdfTextBlock block)
    {
        var text = block.Text.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var compact = text.Replace(" ", string.Empty).Replace("\n", string.Empty);
        if (compact.Length > 160)
        {
            return false;
        }

        const string mathSpecificSymbols = "∑∫∇√≈≤≥≠≃≅∞∈∉∀∃∂∆∏∝⊂⊃⊆⊇∪∩∧∨⊕⊗";
        const string genericMathMarkers = "()[]{}=+-*/^_|<>%";
        var mathSpecificCount = compact.Count(ch => mathSpecificSymbols.Contains(ch));
        var genericMarkerCount = compact.Count(ch => genericMathMarkers.Contains(ch));
        var asciiLetterCount = compact.Count(ch => ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z');
        var digitCount = compact.Count(char.IsDigit);
        var letterOrDigitCount = compact.Count(char.IsLetterOrDigit);
        var symbolRatio = compact.Length == 0 ? 0 : 1 - (letterOrDigitCount / (double)compact.Length);

        var hasLatexCommand = System.Text.RegularExpressions.Regex.IsMatch(text, @"\\[A-Za-z]+(?:\{|\s|$)");
        var hasMathKeyword = System.Text.RegularExpressions.Regex.IsMatch(text, @"\b(arg\s*min|arg\s*max|min|max|lim|sin|cos|tan|log|ln|exp)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var hasMembershipPattern = System.Text.RegularExpressions.Regex.IsMatch(text, @"\b[A-Za-z]\s*[∈∉]\s*[A-Za-z]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var hasSubscriptOrSuperscript = System.Text.RegularExpressions.Regex.IsMatch(text, @"\b[A-Za-z]+(?:_\{?\w+\}?|\^\{?[-+]?\w+\}?)");
        var hasSimpleVariableEquation = System.Text.RegularExpressions.Regex.IsMatch(text, @"\b[A-Za-z]\s*[=<>]\s*[-+]?(?:[A-Za-z0-9]|\([^)]+\))");
        var hasBalancedGrouping = (text.Count(c => c == '(') == text.Count(c => c == ')')) &&
                                  (text.Count(c => c == '[') == text.Count(c => c == ']')) &&
                                  (text.Count(c => c == '{') == text.Count(c => c == '}'));
        var symbolHeavyShortBlock = compact.Length <= 40 && genericMarkerCount >= 2 && symbolRatio >= 0.45;
        var asciiMathExpression = asciiLetterCount > 0 && digitCount > 0 && genericMarkerCount >= 2 && hasBalancedGrouping;

        return hasLatexCommand ||
               hasMembershipPattern ||
               hasSubscriptOrSuperscript ||
               (mathSpecificCount > 0 && genericMarkerCount >= 2) ||
               (hasMathKeyword && genericMarkerCount >= 1) ||
               (hasSimpleVariableEquation && asciiLetterCount <= 12) ||
               symbolHeavyShortBlock ||
               asciiMathExpression;
    }

    // ========================================================================
    // 边缘噪声过滤
    // ========================================================================

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

    // ========================================================================
    // OCR 坐标转换
    // ========================================================================

    /// <summary>
    /// 将 OCR 文本块转换为 PDF 文本块。
    /// 注意坐标系转换：OCR 使用左上角为原点的坐标系，PDF 使用左下角为原点的坐标系。
    /// </summary>
    /// <param name="block">OCR 检测到的文本块</param>
    /// <param name="pageHeight">PDF 页面高度（用于坐标转换）</param>
    /// <returns>转换后的 PDF 文本块</returns>
    private static PdfTextBlock ToPdfTextBlock(OcrTextBlock block, double pageWidth, double pageHeight)
    {
        // OCR 坐标系：原点在左上角，Y 向下增长
        // PDF 坐标系：原点在左下角，Y 向上增长
        // 转换公式：pdfY = pageHeight - ocrY
        var pdfTop = pageHeight - block.Top;
        var pdfBottom = pageHeight - (block.Top + block.Height);

        // 确保坐标有效（防止负数或超出页面）
        pdfTop = Math.Clamp(pdfTop, 0, pageHeight);
        pdfBottom = Math.Clamp(pdfBottom, 0, pageHeight);

        // 确保 Top > Bottom（PDF 坐标系中）
        if (pdfTop < pdfBottom)
        {
            (pdfTop, pdfBottom) = (pdfBottom, pdfTop);
        }

        var pdfBlock = new PdfTextBlock(
            block.Text,
            new PdfRect(
                Math.Max(0, block.Left),
                Math.Max(0, block.Left + block.Width),
                pdfTop,
                pdfBottom),
            Math.Max(10, block.Height),
            XParagraphAlignment.Left,
            new PdfTextStyle(PdfSharpFontResolver.DefaultFontFamily, Math.Max(10, block.Height * 0.9), XColors.Black, false, false),
            PdfBlockType.Normal,
            PdfTextSource.Ocr);

        return RefineBlockRoles([pdfBlock], pageWidth, pageHeight)[0];
    }

    // ========================================================================
    // 字体样式提取（优化）
    // ========================================================================

    private static PdfTextStyle GetTextStyle(IReadOnlyList<Word> words)
    {
        var letters = words.SelectMany(word => word.Letters).ToList();
        if (letters.Count == 0)
        {
            return new PdfTextStyle(PdfSharpFontResolver.DefaultFontFamily, 0, XColors.Black, false, false);
        }

        var fontFamily = letters
            .Select(letter => NormalizePdfFontFamily(letter.FontName))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .Select(group => group.Key)
            .FirstOrDefault() ?? PdfSharpFontResolver.DefaultFontFamily;

        var pointSize = letters
            .Select(letter => letter.PointSize > 0 ? letter.PointSize : letter.FontSize)
            .Where(size => size > 0)
            .DefaultIfEmpty(0)
            .Average();

        var color = letters
            .Select(letter => ConvertColor(letter.Color))
            .GroupBy(colorValue => $"{colorValue.A}-{colorValue.R}-{colorValue.G}-{colorValue.B}")
            .OrderByDescending(group => group.Count())
            .Select(group => group.First())
            .FirstOrDefault();

        // 提取字体粗细和样式信息
        var isBold = letters.Count(l => IsBoldFont(l.FontName)) > letters.Count / 2;
        var isItalic = letters.Count(l => IsItalicFont(l.FontName)) > letters.Count / 2;

        return new PdfTextStyle(fontFamily, pointSize, color, isBold, isItalic);
    }

    private static bool IsBoldFont(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return false;
        }

        var normalized = fontName.ToLowerInvariant();
        return normalized.Contains("bold") ||
               normalized.Contains("heavy") ||
               normalized.Contains("black") ||
               normalized.Contains("semibold") ||
               normalized.Contains("demibold") ||
               normalized.Contains("extrabold") ||
               normalized.Contains("ultrabold") ||
               normalized.Contains("700") ||
               normalized.Contains("800") ||
               normalized.Contains("900");
    }

    private static bool IsItalicFont(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return false;
        }

        var normalized = fontName.ToLowerInvariant();
        return normalized.Contains("italic") ||
               normalized.Contains("oblique") ||
               normalized.Contains("slanted") ||
               normalized.Contains("inclined");
    }

    // ========================================================================
    // 字体解析和映射（优化）
    // ========================================================================

    private static string ResolvePreferredFontFamily(string blockFontFamily, string configuredFontFamily)
    {
        var normalizedBlockFont = NormalizePdfFontFamily(blockFontFamily);
        if (!string.IsNullOrWhiteSpace(normalizedBlockFont))
        {
            return normalizedBlockFont;
        }

        return string.IsNullOrWhiteSpace(configuredFontFamily)
            ? PdfSharpFontResolver.DefaultFontFamily
            : configuredFontFamily;
    }

    private static string NormalizePdfFontFamily(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return string.Empty;
        }

        var normalized = fontName.Trim();
        var plusIndex = normalized.IndexOf('+');
        if (plusIndex > 0 && plusIndex <= 8)
        {
            normalized = normalized[(plusIndex + 1)..];
        }

        // 去除样式后缀以获取纯字体名
        normalized = normalized.Replace('-', ' ')
            .Replace('_', ' ')
            .Replace(',', ' ')
            .Trim();

        // 去除常见的样式描述词以获取基础字体名
        var baseName = System.Text.RegularExpressions.Regex.Replace(
            normalized,
            @"\b(Bold|Italic|Oblique|Light|Medium|Regular|Semibold|Demibold|Condensed|Expanded|Thin|Heavy|Black|ExtraBold|UltraBold|ExtraLight|UltraLight)\b",
            string.Empty,
            System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();

        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = normalized;
        }

        // 衬线字体族映射
        if (baseName.Contains("Times New Roman", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("TimesRoman", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Times", StringComparison.OrdinalIgnoreCase))
        {
            return "Times New Roman";
        }

        if (baseName.Contains("Georgia", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Cambria", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Palatino", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Book Antiqua", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Garamond", StringComparison.OrdinalIgnoreCase))
        {
            return "Times New Roman";
        }

        // 无衬线字体族映射
        if (baseName.Contains("Arial", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Helvetica", StringComparison.OrdinalIgnoreCase))
        {
            return "Arial";
        }

        if (baseName.Contains("Calibri", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Segoe", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Verdana", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Tahoma", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Trebuchet", StringComparison.OrdinalIgnoreCase))
        {
            return "Arial";
        }

        // 等宽字体族映射
        if (baseName.Contains("Courier", StringComparison.OrdinalIgnoreCase))
        {
            return "Courier New";
        }

        if (baseName.Contains("Consolas", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Menlo", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Monaco", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Fira Code", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Source Code", StringComparison.OrdinalIgnoreCase))
        {
            return "Courier New";
        }

        // CJK 字体族映射
        if (baseName.Contains("SimSun", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Song", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("宋体", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Songti", StringComparison.OrdinalIgnoreCase))
        {
            return "SimSun";
        }

        if (baseName.Contains("SimHei", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("Hei", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("黑体", StringComparison.OrdinalIgnoreCase))
        {
            return "SimHei";
        }

        if (baseName.Contains("YaHei", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("微软雅黑", StringComparison.OrdinalIgnoreCase))
        {
            return "Microsoft YaHei";
        }

        if (baseName.Contains("Deng", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("等线", StringComparison.OrdinalIgnoreCase))
        {
            return "DengXian";
        }

        if (baseName.Contains("KaiTi", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("楷体", StringComparison.OrdinalIgnoreCase))
        {
            return "DengXian"; // 映射到最近的可用字体
        }

        if (baseName.Contains("FangSong", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("仿宋", StringComparison.OrdinalIgnoreCase))
        {
            return "DengXian"; // 映射到最近的可用字体
        }

        // 日文字体映射
        if (baseName.Contains("Mincho", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("明朝", StringComparison.OrdinalIgnoreCase))
        {
            return "SimSun";
        }

        if (baseName.Contains("Gothic", StringComparison.OrdinalIgnoreCase) ||
            baseName.Contains("ゴシック", StringComparison.OrdinalIgnoreCase))
        {
            return "SimHei";
        }

        // Noto字体系列
        if (baseName.Contains("Noto Sans", StringComparison.OrdinalIgnoreCase))
        {
            return "DengXian";
        }

        if (baseName.Contains("Noto Serif", StringComparison.OrdinalIgnoreCase))
        {
            return "SimSun";
        }

        return baseName;
    }

    // ========================================================================
    // 颜色转换
    // ========================================================================

    private static XColor ConvertColor(IColor? color) =>
        color switch
        {
            RGBColor rgb => XColor.FromArgb(
                255,
                ToByte(rgb.R),
                ToByte(rgb.G),
                ToByte(rgb.B)),
            GrayColor gray => XColor.FromArgb(
                255,
                ToByte(gray.Gray),
                ToByte(gray.Gray),
                ToByte(gray.Gray)),
            CMYKColor cmyk => ConvertCmykColor(cmyk),
            _ => XColors.Black
        };

    private static XColor ConvertCmykColor(CMYKColor cmyk)
    {
        var cyan = Math.Clamp(cmyk.C, 0, 1);
        var magenta = Math.Clamp(cmyk.M, 0, 1);
        var yellow = Math.Clamp(cmyk.Y, 0, 1);
        var key = Math.Clamp(cmyk.K, 0, 1);
        var red = (1 - cyan) * (1 - key);
        var green = (1 - magenta) * (1 - key);
        var blue = (1 - yellow) * (1 - key);
        return XColor.FromArgb(255, ToByte(red), ToByte(green), ToByte(blue));
    }

    private static byte ToByte(double component) =>
        (byte)Math.Clamp((int)Math.Round(component * 255), 0, 255);

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial)
    {
        progressService.Publish(Path.GetFileName(sourcePath), partial);
        return Task.CompletedTask;
    }

    // ========================================================================
    // 数据模型
    // ========================================================================

    /// <summary>
    /// 文本块类型枚举。
    /// </summary>
    private enum PdfBlockType
    {
        Normal,
        ListItem,
        Code,
        TableRow,
        Title,
        Caption,
        HeaderFooter,
        Footnote
    }

    private enum PdfTextSource
    {
        Native,
        Ocr
    }

    private sealed record PdfRect(double Left, double Right, double Top, double Bottom)
    {
        public double Width => Right - Left;
        public double Height => Top - Bottom;
        public double CenterY => (Top + Bottom) / 2;
    }

    private sealed record PdfTextBlock(
        string Text,
        PdfRect Rect,
        double LineHeight,
        XParagraphAlignment Alignment,
        PdfTextStyle Style,
        PdfBlockType BlockType = PdfBlockType.Normal,
        PdfTextSource Source = PdfTextSource.Native)
    {
        public double Left => Rect.Left;
        public double Right => Rect.Right;
        public double Top => Rect.Top;
        public double Bottom => Rect.Bottom;
        public double Width => Rect.Width;
        public double Height => Rect.Height;
        public double CenterY => Rect.CenterY;
        public string FontFamily => Style.FontFamily;
        public double FontPointSize => Style.PointSize;
        public XColor TextColor => Style.Color;

        public PdfTextBlock Merge(string lineText, PdfRect rect, double lineHeight, PdfTextStyle style) =>
            new(
                MergeText(Text, lineText),
                new PdfRect(
                    Math.Min(Rect.Left, rect.Left),
                    Math.Max(Rect.Right, rect.Right),
                Math.Max(Rect.Top, rect.Top),
                Math.Min(Rect.Bottom, rect.Bottom)),
                Math.Max(LineHeight, lineHeight),
                Alignment,
                ChooseStyle(Style, style),
                BlockType,
                Source);

        public PdfTextBlock WithBlockType(PdfBlockType blockType) => this with { BlockType = blockType };

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

        private static PdfTextStyle ChooseStyle(PdfTextStyle current, PdfTextStyle next)
        {
            if (string.IsNullOrWhiteSpace(current.FontFamily) && !string.IsNullOrWhiteSpace(next.FontFamily))
            {
                return next;
            }

            if (current.PointSize <= 0 && next.PointSize > 0)
            {
                return next;
            }

            return current;
        }
    }

    private sealed record PdfTextStyle(string FontFamily, double PointSize, XColor Color, bool IsBold, bool IsItalic);

    private sealed record PdfTextLayout(XFont Font, IReadOnlyList<string> Lines, double LineHeightRatio, bool Fits);

    private sealed record PdfRenderProfile(
        double SideMargin,
        double TopMargin,
        double BottomMargin,
        double InnerHorizontalPadding,
        double InnerVerticalPadding,
        double LineHeightRatio,
        double MinFontSize,
        double MinWidthInEm,
        double MinHeightInEm,
        bool AllowAggressiveExpansion);

    private sealed record PdfRenderPlan(
        XRect Rect,
        XFont Font,
        IReadOnlyList<string> Lines,
        double LineHeightRatio,
        XColor OverlayColor,
        bool UsedEmergencyFallback);
}
