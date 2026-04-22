using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.IO;
using TranslatorApp.Configuration;
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
        ArgumentNullException.ThrowIfNull(context);

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
        var preparedPages = new List<PreparedPdfPage>(inputPdf.PageCount);
        var requestedRange = GetRequestedRange(context.Settings, inputPdf.PageCount);
        var layoutHeuristics = BuildLayoutHeuristics(context.Settings.Translation);
        var minimumNativeTextWords = Math.Clamp(context.Settings.Ocr.MinimumNativeTextWords, 0, 500);
        if (minimumNativeTextWords != context.Settings.Ocr.MinimumNativeTextWords)
        {
            Log($"OCR 原生文本阈值 {context.Settings.Ocr.MinimumNativeTextWords} 超出合理范围，已自动修正为 {minimumNativeTextWords}。");
        }

        for (var pageIndex = 0; pageIndex < inputPdf.PageCount; pageIndex++)
        {
            await context.PauseController.WaitIfPausedAsync(context.CancellationToken);
            context.CancellationToken.ThrowIfCancellationRequested();

            var pigPage = pig.GetPage(pageIndex + 1);
            var blocks = BuildTextBlocks(pigPage, layoutHeuristics);
            var useOcr = ShouldUseOcr(pigPage, blocks, context.Settings.Ocr, minimumNativeTextWords);
            if (useOcr)
            {
                try
                {
                    var sourcePage = inputPdf.Pages[pageIndex];
                    var ocrBlocks = await ocrService.RecognizePdfPageAsync(context.Item.SourcePath, pageIndex, context.Settings.Ocr, context.CancellationToken);
                    if (ocrBlocks.Count > 0)
                    {
                        blocks = ocrBlocks
                            .Where(block => ShouldKeepOcrBlock(block, sourcePage.Width.Point, sourcePage.Height.Point))
                            .Select(block => ToPdfTextBlock(block, sourcePage.Width.Point, sourcePage.Height.Point))
                            .ToList();
                        Log($"PDF 第 {pageIndex + 1} 页已切换到 OCR 模式。");
                    }
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    Log($"PDF 第 {pageIndex + 1} 页 OCR 失败，已回退到原生文本提取：{ex.Message}");
                }
            }

            preparedPages.Add(new PreparedPdfPage(pageIndex, blocks));
        }

        preparedPages = NormalizeBoundaryHyphenation(preparedPages, layoutHeuristics);
        preparedPages = NormalizeContinuationBlocks(preparedPages, layoutHeuristics);

        for (var pageIndex = 0; pageIndex < inputPdf.PageCount; pageIndex++)
        {
            await context.PauseController.WaitIfPausedAsync(context.CancellationToken);
            context.CancellationToken.ThrowIfCancellationRequested();

            var importedPage = outputPdf.AddPage(inputPdf.Pages[pageIndex]);
            if (!IsWithinRequestedRange(pageIndex + 1, requestedRange))
            {
                continue;
            }

            var blocks = preparedPages[pageIndex].Blocks;
            using var graphics = XGraphics.FromPdfPage(importedPage, XGraphicsPdfPageOptions.Append);
            var translatableBlocks = blocks
                .Select((block, index) => new { Block = block, BlockIndex = index })
                .Where(x => !string.IsNullOrWhiteSpace(x.Block.Text))
                .ToList();

            var translationMap = new Dictionary<int, List<PdfTranslatedTarget>>();
            var translationUnits = BuildPdfTranslationUnits(blocks, layoutHeuristics);
            var batchSize = GetBlockTranslationConcurrency(context.Settings);

            for (var batchStart = 0; batchStart < translationUnits.Count; batchStart += batchSize)
            {
                var batch = translationUnits.Skip(batchStart).Take(batchSize).ToList();
                var translatedBatch = await TranslateBatchAsync(
                    batch.Select(x => CreatePdfTranslationBlock(preparedPages, pageIndex, x.ContextBlockIndex, x.SourceText)).ToList(),
                    context.Settings,
                    context.PauseController,
                    partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                    context.CancellationToken);

                for (var batchIndex = 0; batchIndex < batch.Count; batchIndex++)
                {
                    var unit = batch[batchIndex];
                    var translated = translatedBatch[batchIndex] ?? string.Empty;
                    var distributed = DistributeTranslationAcrossBlocks(unit, translated);
                    foreach (var item in distributed)
                    {
                        if (!translationMap.TryGetValue(item.BlockIndex, out var items))
                        {
                            items = [];
                            translationMap[item.BlockIndex] = items;
                        }

                        items.Add(item);
                    }
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

                    var translated = BuildTranslatedBlockText(translationMap.GetValueOrDefault(entry.BlockIndex));
                    bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", entry.Block.Text, translated));
                    DrawTranslatedBlock(graphics, importedPage.Width.Point, importedPage.Height.Point, entry.Block, translated, context.Settings);
                }
                catch (Exception ex)
                {
                    Log($"PDF 第 {pageIndex + 1} 页文本块 {entry.BlockIndex + 1} 处理失败，已跳过该块：{ex.Message}");
                    bilingualSegments.Add(new BilingualSegment($"PDF 第 {pageIndex + 1} 页", entry.Block.Text, entry.Block.Text));
                }
            }

            var rangeLength = requestedRange.End - requestedRange.Start + 1;
            var completedInRange = pageIndex - requestedRange.Start + 2;
            var progress = (int)Math.Round(completedInRange * 100d / Math.Max(1, rangeLength));
            await context.ReportProgressAsync(progress, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}（范围 {requestedRange.Start}-{requestedRange.End}），文本块 {blocks.Count}");
            await context.SaveCheckpointAsync(completedInRange, 0, $"PDF 页面 {pageIndex + 1}/{inputPdf.PageCount}");
        }

        outputPdf.Save(outputPath);

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static TranslationBlock CreatePdfTranslationBlock(
        IReadOnlyList<PreparedPdfPage> pages,
        int pageIndex,
        int blockIndex,
        string? overrideText = null)
    {
        var block = pages[pageIndex].Blocks[blockIndex];
        var contextHint = BuildPdfContextHint(pages, pageIndex, blockIndex);
        if (ContainsFormulaContent(block))
        {
            contextHint += "；当前片段包含公式、变量或编号。请翻译可翻译的正文内容，并严格保留公式、变量名、运算符、编号和引用格式。";
        }

        return new TranslationBlock(overrideText ?? block.Text, contextHint, BuildPdfAdditionalRequirements(block));
    }

    private static string BuildPdfAdditionalRequirements(PdfTextBlock block)
    {
        var requirements = new List<string>();

        if (block.Region != PdfBlockRegion.Body)
        {
            requirements.Add($"页面区域：{DescribeRegion(block.Region)}。请保持该区域常见的表达方式和结构，不要把它改写成普通正文。");
        }

        switch (block.BlockType)
        {
            case PdfBlockType.Title:
                requirements.Add("类型：标题。请保持标题风格，译文尽量简洁，不要擅自补充说明或改写成正文语气。");
                break;
            case PdfBlockType.Caption:
                requirements.Add("类型：图注。请保留图号、表号、子图编号（如(a)/(b)）、单位、变量名和引用标记；语言保持简洁，不要扩写。");
                break;
            case PdfBlockType.TableRow:
                requirements.Add("类型：表格。请严格保持列顺序、数字、单位、百分比、正负号和变量符号，不要把多列内容合并成解释性整句。");
                break;
            case PdfBlockType.ListItem:
                requirements.Add("类型：列表。请保留项目符号、编号层级和换行结构，不要把多个条目合并成一段。");
                break;
            case PdfBlockType.Code:
                requirements.Add("类型：代码。只翻译自然语言说明，保留代码、路径、API 名称、命令行参数和符号原样。");
                break;
            case PdfBlockType.Footnote:
                requirements.Add("类型：脚注。请保留脚注编号/符号与引用关系，语言可简洁但不要漏译。");
                break;
            case PdfBlockType.HeaderFooter:
                requirements.Add("类型：页眉页脚。请保留页码、编号、期刊信息和短标题的结构，不要扩写。");
                break;
        }

        if (ContainsFormulaContent(block))
        {
            requirements.Add("当前片段含公式或变量。请仅翻译自然语言部分，严格保留公式、变量名、上下标、编号、引用、单位和符号。");
        }

        if (block.Source == PdfTextSource.Ocr)
        {
            requirements.Add("当前片段来自 OCR。请结合上下文纠正明显断词，但不要臆造原文中不存在的数字、公式或专有名词。");
        }

        if (LooksLikeReferenceEntry(block.Text))
        {
            requirements.Add("类型：参考文献。请保留作者名、年份、期刊/会议名称、卷期、页码、DOI、URL、arXiv 编号和编号样式；仅翻译可翻译的说明性文字，不要改写引用格式。");
        }

        return string.Join("\n", requirements);
    }

    private static List<PdfTranslationUnit> BuildPdfTranslationUnits(IReadOnlyList<PdfTextBlock> blocks, PdfLayoutHeuristics heuristics)
    {
        var units = new List<PdfTranslationUnit>();

        for (var index = 0; index < blocks.Count; index++)
        {
            var block = blocks[index];
            if (string.IsNullOrWhiteSpace(block.Text) || IsFormulaLikeBlock(block))
            {
                continue;
            }

            if (block.BlockType == PdfBlockType.TableRow)
            {
                var cellTexts = SplitTableRowSegments(block.Text);
                if (cellTexts.Count > 1)
                {
                    for (var cellIndex = 0; cellIndex < cellTexts.Count; cellIndex++)
                    {
                        units.Add(new PdfTranslationUnit(
                            [new PdfTranslationTarget(index, cellIndex, "  ", cellTexts[cellIndex])],
                            index,
                            cellTexts[cellIndex]));
                    }

                    continue;
                }
            }

            if (!CanBeGroupedForParagraphTranslation(block))
            {
                units.Add(new PdfTranslationUnit(
                    [new PdfTranslationTarget(index, 0, " ", block.Text)],
                    index,
                    block.Text));
                continue;
            }

            var groupedIndices = new List<int> { index };
            var groupedTexts = new List<string> { block.Text };
            var builder = new System.Text.StringBuilder(block.Text.Trim());
            while (!EndsWithSentencePunctuation(builder.ToString()) && index + 1 < blocks.Count)
            {
                var nextIndex = index + 1;
                var nextBlock = blocks[nextIndex];
                if (!CanBeGroupedWithParagraph(blocks[groupedIndices[^1]], nextBlock, heuristics))
                {
                    break;
                }

                AppendGroupedBlockText(builder, nextBlock.Text);
                groupedIndices.Add(nextIndex);
                groupedTexts.Add(nextBlock.Text);
                index = nextIndex;
            }

            units.Add(new PdfTranslationUnit(
                groupedIndices.Select(blockIndex => new PdfTranslationTarget(blockIndex, 0, " ", blocks[blockIndex].Text)).ToList(),
                groupedIndices[0],
                builder.ToString().Trim()));
        }

        return units;
    }

    private static List<PdfTranslatedTarget> DistributeTranslationAcrossBlocks(PdfTranslationUnit unit, string translated)
    {
        if (unit.Targets.Count == 1)
        {
            var target = unit.Targets[0];
            return [new PdfTranslatedTarget(target.BlockIndex, target.Order, translated, target.Joiner)];
        }

        var weights = unit.Targets.Select(target => Math.Max(1, target.OriginalText.Trim().Length)).ToList();
        var segments = TextDistributionHelper.Distribute(translated, weights);
        var result = new List<PdfTranslatedTarget>(unit.Targets.Count);
        for (var i = 0; i < unit.Targets.Count; i++)
        {
            var target = unit.Targets[i];
            result.Add(new PdfTranslatedTarget(
                target.BlockIndex,
                target.Order,
                i < segments.Count ? segments[i].Trim() : string.Empty,
                target.Joiner));
        }

        return result;
    }

    private static IReadOnlyList<string> SplitTableRowSegments(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return [];
        }

        var segments = System.Text.RegularExpressions.Regex
            .Split(text.Trim(), @"(?:\t+|\s{2,}|(?<=\S)\s\|\s(?=\S))")
            .Select(segment => segment.Trim())
            .Where(segment => !string.IsNullOrWhiteSpace(segment))
            .ToList();

        return segments.Count > 1 ? segments : [text.Trim()];
    }

    private static string BuildTranslatedBlockText(IReadOnlyList<PdfTranslatedTarget>? items)
    {
        if (items is null || items.Count == 0)
        {
            return string.Empty;
        }

        var ordered = items.OrderBy(item => item.Order).ToList();
        if (ordered.Count == 1)
        {
            return ordered[0].Text;
        }

        var builder = new System.Text.StringBuilder();
        for (var i = 0; i < ordered.Count; i++)
        {
            if (i > 0)
            {
                builder.Append(ordered[i - 1].Joiner);
            }

            builder.Append(ordered[i].Text);
        }

        return builder.ToString().Trim();
    }

    private static bool CanBeGroupedForParagraphTranslation(PdfTextBlock block) =>
        block.BlockType == PdfBlockType.Normal &&
        block.Region == PdfBlockRegion.Body &&
        !LooksLikeReferenceEntry(block.Text) &&
        CanUseAsTranslationContext(block);

    private static bool CanBeGroupedWithParagraph(PdfTextBlock previous, PdfTextBlock current, PdfLayoutHeuristics heuristics)
    {
        if (!CanBeGroupedForParagraphTranslation(current))
        {
            return false;
        }

        if (previous.Region != current.Region)
        {
            return false;
        }

        var lineHeight = Math.Max(previous.LineHeight, current.LineHeight);
        var verticalGap = previous.Bottom - current.Top;
        if (verticalGap < -2 || verticalGap > lineHeight * heuristics.ParagraphGroupingMaxVerticalGapRatio)
        {
            return false;
        }

        var previousText = previous.Text.TrimEnd();
        var currentText = current.Text.TrimStart();
        if (EndsWithSentencePunctuation(previousText))
        {
            return false;
        }

        if (EndsWithHyphen(previousText))
        {
            return true;
        }

        if (StartsWithContinuationCue(currentText) || StartsWithLowercaseWord(currentText))
        {
            return true;
        }

        return HasParagraphContinuationGeometry(previous, current, lineHeight, allowLooseWrap: true, heuristics);
    }

    private static void AppendGroupedBlockText(System.Text.StringBuilder builder, string nextText)
    {
        var current = builder.ToString().TrimEnd();
        var next = (nextText ?? string.Empty).TrimStart();
        builder.Clear();
        if (string.IsNullOrWhiteSpace(current))
        {
            builder.Append(next);
            return;
        }

        if (string.IsNullOrWhiteSpace(next))
        {
            builder.Append(current);
            return;
        }

        if (EndsWithHyphen(current))
        {
            builder.Append(current.AsSpan(0, TrimTrailingHyphenLength(current)));
            builder.Append(next);
            return;
        }

        builder.Append(current);
        builder.Append(' ');
        builder.Append(next);
    }

    private static string BuildPdfContextHint(
        IReadOnlyList<PreparedPdfPage> pages,
        int pageIndex,
        int blockIndex)
    {
        var block = pages[pageIndex].Blocks[blockIndex];
        var previous = FindNeighborBlockText(pages, pageIndex, blockIndex, searchBackward: true, block.Region);
        var next = FindNeighborBlockText(pages, pageIndex, blockIndex, searchBackward: false, block.Region);
        var location = $"PDF 第 {pageIndex + 1} 页文本块 {blockIndex + 1}";
        var hints = new List<string> { location };

        if (block.Region != PdfBlockRegion.Body)
        {
            hints.Add($"区域：{DescribeRegion(block.Region)}");
        }

        if (!string.IsNullOrWhiteSpace(previous))
        {
            hints.Add($"前文：{previous}");
        }

        if (!string.IsNullOrWhiteSpace(next))
        {
            hints.Add($"后文：{next}");
        }

        if (LooksLikeCrossPageContinuation(pages, pageIndex, blockIndex))
        {
            hints.Add("当前文本很可能是跨页续句。请结合前后文完整翻译当前片段，但不要把前后文一并输出。");
        }

        return string.Join("；", hints);
    }

    private static List<PreparedPdfPage> NormalizeBoundaryHyphenation(IReadOnlyList<PreparedPdfPage> pages, PdfLayoutHeuristics heuristics)
    {
        if (pages.Count == 0)
        {
            return pages.ToList();
        }

        var normalizedPages = pages
            .Select(page => new PreparedPdfPage(page.PageIndex, page.Blocks.ToList()))
            .ToList();

        for (var pageIndex = 0; pageIndex < normalizedPages.Count; pageIndex++)
        {
            var currentBlocks = normalizedPages[pageIndex].Blocks.ToList();
            for (var blockIndex = 0; blockIndex < currentBlocks.Count; blockIndex++)
            {
                var currentText = currentBlocks[blockIndex].Text;
                if (!EndsWithHyphen(currentText))
                {
                    continue;
                }

                if (!TryFindNextHyphenContinuationBlock(normalizedPages, pageIndex, blockIndex, heuristics, out var nextPageIndex, out var nextBlockIndex))
                {
                    continue;
                }

                var nextBlocks = normalizedPages[nextPageIndex].Blocks.ToList();
                if (!TryMergeBoundaryHyphenatedWord(currentText, nextBlocks[nextBlockIndex].Text, out var mergedCurrent, out var trimmedNext))
                {
                    continue;
                }

                currentBlocks[blockIndex] = currentBlocks[blockIndex].WithText(mergedCurrent);
                normalizedPages[pageIndex] = normalizedPages[pageIndex] with { Blocks = currentBlocks };
                nextBlocks[nextBlockIndex] = nextBlocks[nextBlockIndex].WithText(trimmedNext);
                normalizedPages[nextPageIndex] = normalizedPages[nextPageIndex] with { Blocks = nextBlocks };
            }
        }

        return normalizedPages;
    }

    private static List<PreparedPdfPage> NormalizeContinuationBlocks(IReadOnlyList<PreparedPdfPage> pages, PdfLayoutHeuristics heuristics)
    {
        if (pages.Count == 0)
        {
            return pages.ToList();
        }

        return pages
            .Select(page => new PreparedPdfPage(page.PageIndex, MergeContinuationBlocks(page.Blocks, heuristics)))
            .ToList();
    }

    private static IReadOnlyList<PdfTextBlock> MergeContinuationBlocks(IReadOnlyList<PdfTextBlock> blocks, PdfLayoutHeuristics heuristics)
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

            if (ShouldMergeContinuationBlocks(previous, current, heuristics))
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

    private static bool ShouldMergeContinuationBlocks(PdfTextBlock previous, PdfTextBlock current, PdfLayoutHeuristics heuristics)
    {
        if (IsFormulaLikeBlock(previous) || IsFormulaLikeBlock(current))
        {
            return false;
        }

        if (!CanUseAsTranslationContext(previous) || !CanUseAsTranslationContext(current))
        {
            return false;
        }

        if (previous.BlockType != PdfBlockType.Normal || current.BlockType != PdfBlockType.Normal)
        {
            return false;
        }

        if (previous.Region != current.Region)
        {
            return false;
        }

        var lineHeight = Math.Max(previous.LineHeight, current.LineHeight);
        var verticalGap = previous.Bottom - current.Top;
        if (verticalGap < -2 || verticalGap > lineHeight * heuristics.ContinuationMergeMaxVerticalGapRatio)
        {
            return false;
        }

        var sameColumn = HasParagraphContinuationGeometry(previous, current, lineHeight, allowLooseWrap: true, heuristics);
        if (!sameColumn)
        {
            return false;
        }

        var previousTrimmed = previous.Text.TrimEnd();
        var currentTrimmed = current.Text.TrimStart();
        var previousEndsSentence = EndsWithSentencePunctuation(previousTrimmed);
        var currentStartsSentence = StartsWithSentenceCandidate(currentTrimmed);
        var previousParagraphSized = previous.Width > lineHeight * 10;
        var currentParagraphSized = current.Width > lineHeight * 10;
        if (previousEndsSentence && currentStartsSentence)
        {
            return false;
        }

        if (EndsWithHyphen(previousTrimmed) && StartsWithLowercaseWord(currentTrimmed))
        {
            return true;
        }

        var currentStartsContinuation = StartsWithContinuationCue(currentTrimmed) || StartsWithLowercaseWord(currentTrimmed);
        if (!sameColumn &&
            !previousEndsSentence &&
            currentStartsContinuation &&
            previousParagraphSized &&
            currentParagraphSized &&
            current.Left <= previous.Right + Math.Max(42, lineHeight * heuristics.ParagraphLooseWrapForwardRatio) &&
            current.Right >= previous.Left + Math.Max(30, lineHeight * heuristics.ParagraphLooseWrapBackwardRatio))
        {
            return true;
        }

        return !previousEndsSentence && currentStartsContinuation;
    }

    private static bool TryMergeBoundaryHyphenatedWord(string currentText, string nextText, out string mergedCurrent, out string trimmedNext)
    {
        mergedCurrent = currentText;
        trimmedNext = nextText;

        var currentTrimIndex = TrimTrailingHyphenLength(currentText);
        if (currentTrimIndex >= currentText.Length)
        {
            return false;
        }

        var currentPrefix = currentText[..currentTrimIndex];
        var trailingLetters = GetTrailingAsciiLetters(currentPrefix);
        if (trailingLetters.Length < 2)
        {
            return false;
        }

        var nextStart = GetLeadingLowercaseAsciiWordLength(nextText);
        if (nextStart <= 0)
        {
            return false;
        }

        var carried = nextText[..nextStart];
        mergedCurrent = currentPrefix + carried;
        trimmedNext = nextText[nextStart..].TrimStart();
        return true;
    }

    private static bool TryFindNextHyphenContinuationBlock(
        IReadOnlyList<PreparedPdfPage> pages,
        int pageIndex,
        int blockIndex,
        PdfLayoutHeuristics heuristics,
        out int nextPageIndex,
        out int nextBlockIndex)
    {
        var sourceBlock = pages[pageIndex].Blocks[blockIndex];

        for (var currentPage = pageIndex; currentPage < pages.Count; currentPage++)
        {
            var startIndex = currentPage == pageIndex ? blockIndex + 1 : 0;
            for (var currentBlock = startIndex; currentBlock < pages[currentPage].Blocks.Count; currentBlock++)
            {
                var candidate = pages[currentPage].Blocks[currentBlock];
                if (!CanUseAsTranslationContext(candidate))
                {
                    continue;
                }

                if (IsLikelyHyphenContinuationTarget(sourceBlock, candidate, samePage: currentPage == pageIndex, heuristics))
                {
                    nextPageIndex = currentPage;
                    nextBlockIndex = currentBlock;
                    return true;
                }

                nextPageIndex = -1;
                nextBlockIndex = -1;
                return false;
            }
        }

        nextPageIndex = -1;
        nextBlockIndex = -1;
        return false;
    }

    private static bool IsLikelyHyphenContinuationTarget(PdfTextBlock source, PdfTextBlock candidate, bool samePage, PdfLayoutHeuristics heuristics)
    {
        if (!StartsWithLowercaseWord(candidate.Text))
        {
            return false;
        }

        var lineHeight = Math.Max(source.LineHeight, candidate.LineHeight);
        var sameColumn = HasParagraphContinuationGeometry(source, candidate, lineHeight, allowLooseWrap: true, heuristics);

        if (samePage)
        {
            var verticalGap = source.Bottom - candidate.Top;
            var gapLooksContinuous = verticalGap >= -2 && verticalGap < lineHeight * 2.8;
            if (!gapLooksContinuous)
            {
                return false;
            }

            if (sameColumn)
            {
                return true;
            }

            return candidate.Left <= source.Right + Math.Max(48, lineHeight * (heuristics.ParagraphLooseWrapForwardRatio + 0.5)) &&
                   candidate.Right >= source.Left + Math.Max(30, lineHeight * heuristics.ParagraphLooseWrapBackwardRatio);
        }

        return sameColumn || source.Width > candidate.Width * 0.7 || candidate.Width > source.Width * 0.7;
    }

    private static bool HasParagraphContinuationGeometry(
        PdfTextBlock previous,
        PdfTextBlock current,
        double lineHeight,
        bool allowLooseWrap,
        PdfLayoutHeuristics heuristics)
    {
        var leftAligned = Math.Abs(previous.Left - current.Left) < Math.Max(18, lineHeight * heuristics.ParagraphLeftAlignToleranceRatio);
        var rightAligned = Math.Abs(previous.Right - current.Right) < Math.Max(26, lineHeight * heuristics.ParagraphRightAlignToleranceRatio);
        var overlap = Math.Min(previous.Right, current.Right) - Math.Max(previous.Left, current.Left);
        var overlapRatio = overlap / Math.Max(1, Math.Min(previous.Width, current.Width));
        if (overlapRatio > heuristics.ParagraphOverlapThreshold || leftAligned || rightAligned)
        {
            return true;
        }

        if (!allowLooseWrap)
        {
            return false;
        }

        // 对论文正文的换行续写更宽松：上一行可能偏短，下一行重新回到左边界。
        var horizontalGap = current.Left - previous.Right;
        var horizontallyNearby = horizontalGap < Math.Max(28, lineHeight * heuristics.ParagraphHorizontalGapRatio);
        var rangesStillRelated = current.Right > previous.Left + Math.Max(24, lineHeight * heuristics.ParagraphRangeRelationRatio) &&
                                 previous.Right > current.Left - Math.Max(24, lineHeight * heuristics.ParagraphRangeRelationRatio);
        var paragraphSized = previous.Width > lineHeight * heuristics.ParagraphMinWidthRatio &&
                             current.Width > lineHeight * heuristics.ParagraphMinWidthRatio;

        if (paragraphSized && horizontallyNearby && rangesStillRelated)
        {
            return true;
        }

        return paragraphSized &&
               current.Left <= previous.Right + Math.Max(42, lineHeight * heuristics.ParagraphLooseWrapForwardRatio) &&
               current.Right >= previous.Left + Math.Max(30, lineHeight * heuristics.ParagraphLooseWrapBackwardRatio);
    }

    private static string FindNeighborBlockText(
        IReadOnlyList<PreparedPdfPage> pages,
        int pageIndex,
        int blockIndex,
        bool searchBackward,
        PdfBlockRegion preferredRegion)
    {
        var sameRegion = SearchNeighborBlockText(pages, pageIndex, blockIndex, searchBackward, candidate => candidate.Region == preferredRegion);
        if (!string.IsNullOrWhiteSpace(sameRegion))
        {
            return sameRegion;
        }

        return SearchNeighborBlockText(pages, pageIndex, blockIndex, searchBackward, _ => true);
    }

    private static string SearchNeighborBlockText(
        IReadOnlyList<PreparedPdfPage> pages,
        int pageIndex,
        int blockIndex,
        bool searchBackward,
        Func<PdfTextBlock, bool> predicate)
    {
        if (searchBackward)
        {
            for (var currentPage = pageIndex; currentPage >= 0; currentPage--)
            {
                var startIndex = currentPage == pageIndex ? blockIndex - 1 : pages[currentPage].Blocks.Count - 1;
                for (var currentBlock = startIndex; currentBlock >= 0; currentBlock--)
                {
                    var candidate = pages[currentPage].Blocks[currentBlock];
                    if (CanUseAsTranslationContext(candidate) && predicate(candidate))
                    {
                        return BuildContextPreview(candidate.Text);
                    }
                }
            }

            return string.Empty;
        }

        for (var currentPage = pageIndex; currentPage < pages.Count; currentPage++)
        {
            var startIndex = currentPage == pageIndex ? blockIndex + 1 : 0;
            for (var currentBlock = startIndex; currentBlock < pages[currentPage].Blocks.Count; currentBlock++)
            {
                var candidate = pages[currentPage].Blocks[currentBlock];
                if (CanUseAsTranslationContext(candidate) && predicate(candidate))
                {
                    return BuildContextPreview(candidate.Text);
                }
            }
        }

        return string.Empty;
    }

    private static bool CanUseAsTranslationContext(PdfTextBlock block) =>
        !string.IsNullOrWhiteSpace(block.Text) &&
        !IsFormulaLikeBlock(block) &&
        !LooksLikeReferenceEntry(block.Text) &&
        block.BlockType is not PdfBlockType.HeaderFooter and not PdfBlockType.Footnote &&
        block.Region is not PdfBlockRegion.Margin;

    private static string DescribeRegion(PdfBlockRegion region) => region switch
    {
        PdfBlockRegion.Body => "正文区",
        PdfBlockRegion.Caption => "图注区",
        PdfBlockRegion.Table => "表格区",
        PdfBlockRegion.Margin => "边注区",
        PdfBlockRegion.HeaderFooter => "页眉页脚区",
        PdfBlockRegion.Footnote => "脚注区",
        _ => "正文区"
    };

    private static string BuildContextPreview(string text)
    {
        var normalized = string.Join(" ", text
            .Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries));
        return normalized.Length <= 180 ? normalized : normalized[..180] + "...";
    }

    private static bool LooksLikeCrossPageContinuation(
        IReadOnlyList<PreparedPdfPage> pages,
        int pageIndex,
        int blockIndex)
    {
        var block = pages[pageIndex].Blocks[blockIndex];
        if (pageIndex > 0 && blockIndex == 0)
        {
            var previousPageTail = FindNeighborBlockText(pages, pageIndex, blockIndex, searchBackward: true, block.Region);
            if (!string.IsNullOrWhiteSpace(previousPageTail) && StartsWithContinuationCue(block.Text))
            {
                return true;
            }
        }

        if (pageIndex < pages.Count - 1 && blockIndex == pages[pageIndex].Blocks.Count - 1)
        {
            var nextPageHead = FindNeighborBlockText(pages, pageIndex, blockIndex, searchBackward: false, block.Region);
            if (!string.IsNullOrWhiteSpace(nextPageHead) && EndsWithContinuationCue(block.Text))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldUseOcr(Page page, IReadOnlyList<PdfTextBlock> blocks, OcrSettings ocrSettings, int minimumNativeTextWords)
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
                               textCoverage < ocrSettings.SparseTextCoverageThreshold &&
                               blocks.Count <= ocrSettings.SparseTextBlockThreshold;
        return sparseNativeText;
    }

    // ========================================================================
    // 文本块检测和边界框计算（优化）
    // ========================================================================

    private static List<PdfTextBlock> BuildTextBlocks(Page page, PdfLayoutHeuristics heuristics)
    {
        var words = page.GetWords()
            .Where(x => !string.IsNullOrWhiteSpace(x.Text))
            .Where(x => !IsLikelyMarginalNoise(x, page.Width, heuristics))
            .OrderByDescending(x => x.BoundingBox.Top)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        // 第一步：检测多列布局
        var columns = DetectColumns(words, page.Width, heuristics);

        // 第二步：在每列内部构建行和段落块
        var allBlocks = new List<PdfTextBlock>();
        foreach (var columnWords in columns)
        {
            var columnBlocks = BuildColumnBlocks(columnWords, page.Width, heuristics);
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
    private static List<List<Word>> DetectColumns(IReadOnlyList<Word> words, double pageWidth, PdfLayoutHeuristics heuristics)
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
        var minColumnGap = pageWidth * heuristics.ColumnGapRatio;

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

            if (leftWords.Count < heuristics.ColumnMinWordsPerSide || rightWords.Count < heuristics.ColumnMinWordsPerSide)
            {
                continue;
            }

            // 检查两侧文本是否都有足够的垂直跨度
            var leftVertSpan = leftWords.Max(w => w.BoundingBox.Top) - leftWords.Min(w => w.BoundingBox.Bottom);
            var rightVertSpan = rightWords.Max(w => w.BoundingBox.Top) - rightWords.Min(w => w.BoundingBox.Bottom);

            if (leftVertSpan > pageContentHeight * heuristics.ColumnMinVerticalSpanRatio &&
                rightVertSpan > pageContentHeight * heuristics.ColumnMinVerticalSpanRatio)
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
    private static List<PdfTextBlock> BuildColumnBlocks(List<Word> words, double pageWidth, PdfLayoutHeuristics heuristics)
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

            if (isCloseLine &&
                hasHorizontalRelation &&
                sameType &&
                ShouldMergeLineIntoBlock(previous, lineText, rect, lineHeight, pageWidth, heuristics))
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

    private static bool ShouldMergeLineIntoBlock(
        PdfTextBlock previous,
        string currentLineText,
        PdfRect currentRect,
        double currentLineHeight,
        double pageWidth,
        PdfLayoutHeuristics heuristics)
    {
        if (previous.BlockType is PdfBlockType.TableRow or PdfBlockType.ListItem or PdfBlockType.Code)
        {
            return false;
        }

        var previousIsReferenceEntry = LooksLikeReferenceEntry(previous.Text);
        var currentStartsReferenceEntry = StartsReferenceEntry(currentLineText);
        if (currentStartsReferenceEntry)
        {
            return false;
        }

        if (previousIsReferenceEntry)
        {
            var referenceLineHeight = Math.Max(previous.LineHeight, currentLineHeight);
            var referenceGap = previous.Bottom - currentRect.Top;
            var horizontallyRelated = Math.Abs(previous.Left - currentRect.Left) < Math.Max(18, referenceLineHeight * 1.2) ||
                                      OverlapRatio(previous.Left, previous.Right, currentRect.Left, currentRect.Right) > 0.45;
            return referenceGap >= -2 &&
                   referenceGap <= referenceLineHeight * Math.Max(1.5, heuristics.LineMergeMaxVerticalGapRatio) &&
                   horizontallyRelated;
        }

        var mergeLineHeight = Math.Max(previous.LineHeight, currentLineHeight);
        var verticalGap = previous.Bottom - currentRect.Top;
        if (verticalGap < -2 || verticalGap > mergeLineHeight * heuristics.LineMergeMaxVerticalGapRatio)
        {
            return false;
        }

        var previousTrimmed = previous.Text.TrimEnd();
        var currentTrimmed = currentLineText.TrimStart();
        var previousEndsSentence = EndsWithSentencePunctuation(previousTrimmed);
        var currentStartsSentence = StartsWithSentenceCandidate(currentTrimmed);
        var alignedLeft = Math.Abs(previous.Left - currentRect.Left) < Math.Max(12, mergeLineHeight);
        var indentChange = currentRect.Left - previous.Left;
        var notableIndent = indentChange > mergeLineHeight * 0.8;
        var previousLooksShort = previous.Width < Math.Max(pageWidth * 0.3, currentRect.Width * 0.82);

        if (previousEndsSentence && currentStartsSentence && (notableIndent || previousLooksShort))
        {
            return false;
        }

        if (previousEndsSentence &&
            alignedLeft &&
            previousLooksShort &&
            verticalGap > mergeLineHeight * 0.2)
        {
            return false;
        }

        return true;
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
            var prevEndsWithHyphen = EndsWithHyphen(prevWord.Text);
            var currStartsWithLowercase = StartsWithLowercaseWord(currWord.Text);

            if (prevEndsWithHyphen && currStartsWithLowercase)
            {
                result = string.Concat(result.AsSpan(0, TrimTrailingHyphenLength(result)), currWord.Text.TrimStart());
            }
            else if (prevEndsWithHyphen)
            {
                result = string.Concat(result.AsSpan(0, TrimTrailingHyphenLength(result)), currWord.Text);
            }
            else
            {
                if (prevEndsWithCjk && currStartsWithCjk)
                {
                    // CJK之间不加空格
                    result += currWord.Text;
                }
                else if (ShouldJoinWithoutSpace(prevWord.Text, currWord.Text))
                {
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
        var yielded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var candidate in SplitFontFamilyCandidates(family))
        {
            if (yielded.Add(candidate))
            {
                yield return candidate;
            }
        }

        if (!string.Equals(family, PdfSharpFontResolver.DefaultFontFamily, StringComparison.OrdinalIgnoreCase))
        {
            if (yielded.Add(PdfSharpFontResolver.DefaultFontFamily))
            {
                yield return PdfSharpFontResolver.DefaultFontFamily;
            }
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

        return AssignRegions(refined, pageWidth, pageHeight);
    }

    private static List<PdfTextBlock> AssignRegions(IReadOnlyList<PdfTextBlock> blocks, double pageWidth, double pageHeight)
    {
        if (blocks.Count == 0)
        {
            return [];
        }

        var regioned = blocks
            .Select(block => block.WithRegion(DetectRegion(block, pageWidth, pageHeight)))
            .ToList();

        for (var i = 0; i < regioned.Count; i++)
        {
            var block = regioned[i];
            if (block.BlockType != PdfBlockType.Caption || block.Region == PdfBlockRegion.Table)
            {
                continue;
            }

            var hasNearbyTable = regioned.Any(other =>
                other.BlockType == PdfBlockType.TableRow &&
                Math.Abs(other.CenterY - block.CenterY) < Math.Max(block.LineHeight, other.LineHeight) * 6 &&
                OverlapRatio(block.Left, block.Right, other.Left, other.Right) > 0.2);
            if (hasNearbyTable)
            {
                regioned[i] = block.WithRegion(PdfBlockRegion.Table);
            }
        }

        return regioned;
    }

    private static PdfBlockRegion DetectRegion(PdfTextBlock block, double pageWidth, double pageHeight)
    {
        if (block.BlockType == PdfBlockType.HeaderFooter)
        {
            return PdfBlockRegion.HeaderFooter;
        }

        if (block.BlockType == PdfBlockType.Footnote)
        {
            return PdfBlockRegion.Footnote;
        }

        if (block.BlockType == PdfBlockType.Caption)
        {
            return PdfBlockRegion.Caption;
        }

        if (block.BlockType == PdfBlockType.TableRow)
        {
            return PdfBlockRegion.Table;
        }

        var centerX = (block.Left + block.Right) / 2;
        var inLeftMargin = centerX < pageWidth * 0.12;
        var inRightMargin = centerX > pageWidth * 0.88;
        var narrowBlock = block.Width < pageWidth * 0.22;
        var shortBlock = block.Text.Trim().Length <= 60;
        var nearVerticalEdge = block.Top > pageHeight * 0.08 && block.Bottom < pageHeight * 0.92;
        if ((inLeftMargin || inRightMargin) && narrowBlock && shortBlock && nearVerticalEdge)
        {
            return PdfBlockRegion.Margin;
        }

        return PdfBlockRegion.Body;
    }

    private static double OverlapRatio(double left1, double right1, double left2, double right2)
    {
        var overlap = Math.Min(right1, right2) - Math.Max(left1, left2);
        if (overlap <= 0)
        {
            return 0;
        }

        var width = Math.Max(1, Math.Min(right1 - left1, right2 - left2));
        return overlap / width;
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

        var previousIsReferenceEntry = LooksLikeReferenceEntry(previous.Text);
        var currentStartsReferenceEntry = StartsReferenceEntry(current.Text);
        if (currentStartsReferenceEntry)
        {
            return false;
        }

        if (previousIsReferenceEntry)
        {
            if (current.BlockType is PdfBlockType.Title or PdfBlockType.Caption or PdfBlockType.HeaderFooter or PdfBlockType.Footnote)
            {
                return false;
            }

            var referenceLineHeight = Math.Max(previous.LineHeight, current.LineHeight);
            var referenceGap = previous.Bottom - current.Top;
            var referenceSameColumn = Math.Abs(previous.Left - current.Left) < Math.Max(18, referenceLineHeight * 1.2) ||
                                      OverlapRatio(previous.Left, previous.Right, current.Left, current.Right) > 0.45;
            return referenceGap >= -2 &&
                   referenceGap < referenceLineHeight * 1.6 &&
                   referenceSameColumn;
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
                   previous.Region == current.Region &&
                   Math.Abs(previous.Left - current.Left) < Math.Max(previous.LineHeight, current.LineHeight) &&
                   Math.Abs(previous.Right - current.Right) < Math.Max(previous.LineHeight * 1.5, current.LineHeight * 1.5) &&
                   previous.Bottom - current.Top < Math.Max(previous.LineHeight, current.LineHeight) * 1.4;
        }

        if (previous.Region != current.Region)
        {
            return false;
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
            return false;
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

        if (sameColumn &&
            gapLooksContinuous &&
            EndsWithHyphen(previous.Text) &&
            StartsWithLowercaseWord(current.Text))
        {
            return true;
        }

        // 缩进变化检查（可能是新段落开始）
        var indentChange = Math.Abs(previous.Left - current.Left);
        var significantIndentChange = indentChange > mergeLineHeight * 1.5 && previous.Width < pageWidth * 0.7;

        // 如果缩进显著变化，可能是新段落，不合并
        if (significantIndentChange && !alignedLeft)
        {
            return false;
        }

        var previousTrimmed = previous.Text.TrimEnd();
        var currentTrimmed = current.Text.TrimStart();
        var previousLooksShort = previous.Width < Math.Max(pageWidth * 0.3, current.Width * 0.82);
        var currentStartsSentence = StartsWithSentenceCandidate(currentTrimmed);

        if (EndsWithSentencePunctuation(previousTrimmed) &&
            currentStartsSentence &&
            (significantIndentChange || previousLooksShort || mergeVerticalGap > mergeLineHeight * 0.25))
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

        return paragraphLike && consistentLineHeight && continuationEvidence;
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

    private static bool StartsWithSentenceCandidate(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var trimmed = text.TrimStart();
        var index = 0;
        while (index < trimmed.Length && "\"'([{".Contains(trimmed[index]))
        {
            index++;
        }

        if (index >= trimmed.Length)
        {
            return false;
        }

        return char.IsUpper(trimmed[index]) || IsCjkCharacter(trimmed[index]);
    }

    private static bool StartsWithLowercaseWord(string text)
    {
        var trimmed = text.TrimStart();
        if (string.IsNullOrEmpty(trimmed))
        {
            return false;
        }

        var index = 0;
        while (index < trimmed.Length && "\"'([{".Contains(trimmed[index]))
        {
            index++;
        }

        return index < trimmed.Length && char.IsLower(trimmed[index]);
    }

    private static bool ShouldJoinWithoutSpace(string previousText, string currentText)
    {
        var previousTrimmed = previousText.TrimEnd();
        var currentTrimmed = currentText.TrimStart();
        if (string.IsNullOrEmpty(previousTrimmed) || string.IsNullOrEmpty(currentTrimmed))
        {
            return false;
        }

        var previousLast = previousTrimmed[^1];
        var currentFirst = currentTrimmed[0];
        if (IsOpeningPunctuation(previousLast) || IsClosingPunctuation(currentFirst))
        {
            return true;
        }

        if (previousLast is '/' or '@' or '#')
        {
            return true;
        }

        if (currentFirst is '/' or '%' or '\'' || currentTrimmed.StartsWith("'s", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static bool IsOpeningPunctuation(char ch) => "([{\"'“‘".Contains(ch);

    private static bool IsClosingPunctuation(char ch) => ".,;:!?)]}%\"'”’".Contains(ch);

    private static string GetTrailingAsciiLetters(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var end = text.Length;
        while (end > 0 && char.IsWhiteSpace(text[end - 1]))
        {
            end--;
        }

        var start = end;
        while (start > 0 && IsAsciiLetter(text[start - 1]))
        {
            start--;
        }

        return text[start..end];
    }

    private static int GetLeadingLowercaseAsciiWordLength(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return 0;
        }

        var trimmed = text.TrimStart();
        var index = 0;
        while (index < trimmed.Length && IsAsciiLower(trimmed[index]))
        {
            index++;
        }

        return index;
    }

    private static bool EndsWithHyphen(string text) => TrimTrailingHyphenLength(text) < text.Length;

    private static int TrimTrailingHyphenLength(string text)
    {
        var end = text.Length;
        while (end > 0 && char.IsWhiteSpace(text[end - 1]))
        {
            end--;
        }

        if (end > 0 && IsHyphenChar(text[end - 1]))
        {
            end--;
            while (end > 0 && char.IsWhiteSpace(text[end - 1]))
            {
                end--;
            }
        }

        return end;
    }

    private static bool IsHyphenChar(char ch) =>
        ch is '-' or '‐' or '‑' or '‒' or '–' or '﹣' or '－' or '\u00AD';

    private static bool IsAsciiLetter(char ch) => ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z';

    private static bool IsAsciiLower(char ch) => ch is >= 'a' and <= 'z';

    private static bool StartsReferenceEntry(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return System.Text.RegularExpressions.Regex.IsMatch(
            text.TrimStart(),
            @"^(?:\[\d{1,4}(?:\s*[-,]\s*\d{1,4})?\]|\(\d{1,4}\)|\d{1,4}[\.\)])\s+");
    }

    private static bool LooksLikeReferenceEntry(string text)
    {
        if (!StartsReferenceEntry(text))
        {
            return false;
        }

        var trimmed = text.Trim();
        if (trimmed.Length < 18)
        {
            return false;
        }

        return System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"\b(19|20)\d{2}[a-z]?\b") ||
               System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"\b(?:doi|arxiv|vol\.?|pp\.?|pages?)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase) ||
               System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"\b(?:Proceedings|Journal|Conference|Transactions)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    // ========================================================================
    // 公式检测
    // ========================================================================

    private static bool IsFormulaLikeBlock(PdfTextBlock block) => AnalyzeFormulaBlock(block.Text).IsPureFormula;

    private static bool ContainsFormulaContent(PdfTextBlock block) => AnalyzeFormulaBlock(block.Text).ContainsFormulaContent;

    private static FormulaBlockAnalysis AnalyzeFormulaBlock(string rawText)
    {
        var text = rawText.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return new FormulaBlockAnalysis(false, false);
        }

        var compact = text.Replace(" ", string.Empty).Replace("\n", string.Empty);
        if (compact.Length > 160)
        {
            return new FormulaBlockAnalysis(false, false);
        }

        const string mathSpecificSymbols = "∑∫∇√≈≤≥≠≃≅∞∈∉∀∃∂∆∏∝⊂⊃⊆⊇∪∩∧∨⊕⊗";
        const string groupingMarkers = "()[]{}";
        const string operatorMarkers = "=+*/^_|<>%";
        const string genericMathMarkers = groupingMarkers + operatorMarkers + "-";
        var mathSpecificCount = compact.Count(ch => mathSpecificSymbols.Contains(ch));
        var groupingMarkerCount = compact.Count(ch => groupingMarkers.Contains(ch));
        var operatorMarkerCount = compact.Count(ch => operatorMarkers.Contains(ch));
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
        var looksLikeCitation = System.Text.RegularExpressions.Regex.IsMatch(
            text,
            @"\b(?:see|table|fig|figure|section|sec|eq|equation|ref|appendix)\b[\s\.:,-]*\(?[\dA-Za-z][\dA-Za-z\.\-]*\)?",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var looksLikeVersion = System.Text.RegularExpressions.Regex.IsMatch(text, @"^[vV]?[\d\.]+[a-z]?$");
        var hasBalancedGrouping = (text.Count(c => c == '(') == text.Count(c => c == ')')) &&
                                  (text.Count(c => c == '[') == text.Count(c => c == ']')) &&
                                  (text.Count(c => c == '{') == text.Count(c => c == '}'));
        var symbolHeavyShortBlock = compact.Length <= 40 && genericMarkerCount >= 2 && symbolRatio >= 0.45;
        var proseWordCount = System.Text.RegularExpressions.Regex.Matches(text, @"\b[A-Za-z]{2,}\b").Count;
        var hasProseIndicators = System.Text.RegularExpressions.Regex.IsMatch(
            text,
            @"\b(?:the|and|for|are|with|this|that|see|table|fig|figure|section|sec|eq|equation|ref|below|above|where|when|given|then|thus|let)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var looksLikeProse = proseWordCount >= 2 || hasProseIndicators;
        var hasSentenceLikeStructure = System.Text.RegularExpressions.Regex.IsMatch(
            text,
            @"\b(?:is|are|was|were|be|can|may|shows?|denotes?|represents?)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var hasDefinitionLeadIn = System.Text.RegularExpressions.Regex.IsMatch(
            text,
            @"\b(?:where|when|if|for|given|let|then|thus)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var hasClausePunctuation = text.Contains(',') || text.Contains('，') || text.Contains(':') || text.Contains('：');
        var containsFormulaSignal =
            hasLatexCommand ||
            hasMembershipPattern ||
            hasSubscriptOrSuperscript ||
            hasSimpleVariableEquation ||
            hasMathKeyword ||
            mathSpecificCount > 0 ||
            symbolHeavyShortBlock ||
            asciiLetterCount > 0 && (operatorMarkerCount >= 1 || groupingMarkerCount >= 2) && digitCount > 0;

        if (looksLikeCitation || looksLikeVersion)
        {
            return new FormulaBlockAnalysis(false, containsFormulaSignal);
        }

        if (compact.Length > 80 && proseWordCount >= 3)
        {
            return new FormulaBlockAnalysis(false, containsFormulaSignal);
        }

        var asciiMathExpression =
            asciiLetterCount > 0 &&
            digitCount > 0 &&
            operatorMarkerCount >= 1 &&
            groupingMarkerCount >= 2 &&
            genericMarkerCount >= 3 &&
            hasBalancedGrouping &&
            !looksLikeProse &&
            symbolRatio >= 0.3 &&
            text.Length <= 60;

        var simpleEquationOnly =
            hasSimpleVariableEquation &&
            asciiLetterCount <= 12 &&
            !looksLikeProse &&
            !hasSentenceLikeStructure &&
            !hasDefinitionLeadIn &&
            text.Length <= 24 &&
            proseWordCount <= 1 &&
            digitCount <= 1 &&
            groupingMarkerCount == 0;

        var proseWithFormula =
            containsFormulaSignal &&
            (looksLikeProse || hasSentenceLikeStructure || hasDefinitionLeadIn) &&
            (proseWordCount >= 2 || hasClausePunctuation || text.Length > 32);

        var pureFormula =
            !proseWithFormula &&
            (hasLatexCommand ||
             hasMembershipPattern ||
             hasSubscriptOrSuperscript ||
             (mathSpecificCount > 0 && genericMarkerCount >= 2) ||
             (hasMathKeyword && genericMarkerCount >= 1 && !looksLikeProse) ||
             simpleEquationOnly ||
             symbolHeavyShortBlock ||
             asciiMathExpression);

        return new FormulaBlockAnalysis(pureFormula, containsFormulaSignal || pureFormula);
    }

    private static PdfLayoutHeuristics BuildLayoutHeuristics(TranslationSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return new PdfLayoutHeuristics(
            settings.PdfColumnGapRatio,
            settings.PdfColumnMinWordsPerSide,
            settings.PdfColumnMinVerticalSpanRatio,
            settings.PdfMarginNoiseSideRatio,
            settings.PdfMarginNoiseVerticalAspectRatio,
            settings.PdfMarginNoiseShortTokenLength,
            settings.PdfParagraphGroupingMaxVerticalGapRatio,
            settings.PdfContinuationMergeMaxVerticalGapRatio,
            settings.PdfLineMergeMaxVerticalGapRatio,
            settings.PdfParagraphLeftAlignToleranceRatio,
            settings.PdfParagraphRightAlignToleranceRatio,
            settings.PdfParagraphOverlapThreshold,
            settings.PdfParagraphHorizontalGapRatio,
            settings.PdfParagraphRangeRelationRatio,
            settings.PdfParagraphMinWidthRatio,
            settings.PdfParagraphLooseWrapForwardRatio,
            settings.PdfParagraphLooseWrapBackwardRatio);
    }

    private readonly record struct FormulaBlockAnalysis(bool IsPureFormula, bool ContainsFormulaContent);
    private readonly record struct PdfLayoutHeuristics(
        double ColumnGapRatio,
        int ColumnMinWordsPerSide,
        double ColumnMinVerticalSpanRatio,
        double MarginNoiseSideRatio,
        double MarginNoiseVerticalAspectRatio,
        int MarginNoiseShortTokenLength,
        double ParagraphGroupingMaxVerticalGapRatio,
        double ContinuationMergeMaxVerticalGapRatio,
        double LineMergeMaxVerticalGapRatio,
        double ParagraphLeftAlignToleranceRatio,
        double ParagraphRightAlignToleranceRatio,
        double ParagraphOverlapThreshold,
        double ParagraphHorizontalGapRatio,
        double ParagraphRangeRelationRatio,
        double ParagraphMinWidthRatio,
        double ParagraphLooseWrapForwardRatio,
        double ParagraphLooseWrapBackwardRatio);

    // ========================================================================
    // 边缘噪声过滤
    // ========================================================================

    private static bool IsLikelyMarginalNoise(Word word, double pageWidth, PdfLayoutHeuristics heuristics)
    {
        var boxWidth = Math.Max(1, word.BoundingBox.Right - word.BoundingBox.Left);
        var boxHeight = Math.Max(1, word.BoundingBox.Top - word.BoundingBox.Bottom);
        var centerX = (word.BoundingBox.Left + word.BoundingBox.Right) / 2;
        var sideRatio = heuristics.MarginNoiseSideRatio;
        var inSideMargin = centerX < pageWidth * sideRatio || centerX > pageWidth * (1 - sideRatio);
        var looksVertical = boxHeight > boxWidth * heuristics.MarginNoiseVerticalAspectRatio;
        var shortToken = word.Text.Trim().Length <= heuristics.MarginNoiseShortTokenLength;
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

        return SplitFontFamilyCandidates(configuredFontFamily).FirstOrDefault()
            ?? PdfSharpFontResolver.DefaultFontFamily;
    }

    private static IReadOnlyList<string> SplitFontFamilyCandidates(string? familyNames)
    {
        if (string.IsNullOrWhiteSpace(familyNames))
        {
            return Array.Empty<string>();
        }

        return familyNames
            .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(candidate => candidate.Trim().Trim('"', '\''))
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizePdfFontFamily(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return string.Empty;
        }

        var normalized = SplitFontFamilyCandidates(fontName).FirstOrDefault() ?? fontName.Trim();
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

    private enum PdfBlockRegion
    {
        Body,
        Caption,
        Table,
        Margin,
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
        PdfTextSource Source = PdfTextSource.Native,
        PdfBlockRegion Region = PdfBlockRegion.Body)
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

        public PdfTextBlock WithRegion(PdfBlockRegion region) => this with { Region = region };

        public PdfTextBlock WithText(string text) => this with { Text = text };

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
                ? string.Concat(currentText.AsSpan(0, TrimTrailingHyphenLength(currentText)), nextLine)
                : $"{currentText} {nextLine}";
        }

        private static bool EndsWithHyphen(string text) => TrimTrailingHyphenLength(text) < text.Length;

        private static int TrimTrailingHyphenLength(string text)
        {
            var end = text.Length;
            while (end > 0 && char.IsWhiteSpace(text[end - 1]))
            {
                end--;
            }

            if (end > 0 && IsHyphenChar(text[end - 1]))
            {
                end--;
                while (end > 0 && char.IsWhiteSpace(text[end - 1]))
                {
                    end--;
                }
            }

            return end;
        }

        private static bool IsHyphenChar(char ch) =>
            ch is '-' or '‐' or '‑' or '‒' or '–' or '﹣' or '－' or '\u00AD';

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

    private sealed record PreparedPdfPage(int PageIndex, IReadOnlyList<PdfTextBlock> Blocks);

    private sealed record PdfTranslationTarget(
        int BlockIndex,
        int Order,
        string Joiner,
        string OriginalText);

    private sealed record PdfTranslatedTarget(
        int BlockIndex,
        int Order,
        string Text,
        string Joiner);

    private sealed record PdfTranslationUnit(
        IReadOnlyList<PdfTranslationTarget> Targets,
        int ContextBlockIndex,
        string SourceText);

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
