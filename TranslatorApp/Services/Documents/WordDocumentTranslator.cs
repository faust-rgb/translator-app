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
        var selectedParagraphs = pageInfos
            .Where(x => IsWithinRequestedRange(x.PageNumber, requestedRange))
            .Select(x => x.Paragraph)
            .ToList();

        var translateFullDocumentExtras = requestedRange.Start == 1 && requestedRange.End >= totalPages;
        var partsToSave = new List<OpenXmlPartRootElement> { mainDocument };
        var units = new List<WordTranslationUnit>();

        units.AddRange(BuildMainUnits(selectedParagraphs));

        if (translateFullDocumentExtras)
        {
            var headerRoots = document.MainDocumentPart?.HeaderParts.Select(x => x.Header).OfType<Header>().Cast<OpenXmlPartRootElement>().ToList() ?? [];
            var footerRoots = document.MainDocumentPart?.FooterParts.Select(x => x.Footer).OfType<Footer>().Cast<OpenXmlPartRootElement>().ToList() ?? [];
            var footnotesRoot = document.MainDocumentPart?.FootnotesPart?.Footnotes;
            var endnotesRoot = document.MainDocumentPart?.EndnotesPart?.Endnotes;
            var commentsRoot = document.MainDocumentPart?.WordprocessingCommentsPart?.Comments;

            partsToSave.AddRange(headerRoots);
            partsToSave.AddRange(footerRoots);
            if (footnotesRoot is not null)
            {
                partsToSave.Add(footnotesRoot);
            }

            if (endnotesRoot is not null)
            {
                partsToSave.Add(endnotesRoot);
            }

            if (commentsRoot is not null)
            {
                partsToSave.Add(commentsRoot);
            }

            units.AddRange(BuildUnitsFromRootElements(headerRoots, WordScope.Header));
            units.AddRange(BuildUnitsFromRootElements(footerRoots, WordScope.Footer));
            if (footnotesRoot is not null)
            {
                units.AddRange(BuildUnitsFromRootElement(footnotesRoot, WordScope.Footnote));
            }

            if (endnotesRoot is not null)
            {
                units.AddRange(BuildUnitsFromRootElement(endnotesRoot, WordScope.Endnote));
            }

            if (commentsRoot is not null)
            {
                units.AddRange(BuildUnitsFromRootElement(commentsRoot, WordScope.Comment));
            }
        }
        else if (document.MainDocumentPart?.HeaderParts.Any() == true ||
                 document.MainDocumentPart?.FooterParts.Any() == true ||
                 document.MainDocumentPart?.FootnotesPart is not null ||
                 document.MainDocumentPart?.EndnotesPart is not null ||
                 document.MainDocumentPart?.WordprocessingCommentsPart is not null)
        {
            Log($"Word 当前按近似分页范围 {requestedRange.Start}-{requestedRange.End} 处理，页眉页脚、脚注尾注和批注仅在全量范围时一并翻译。");
        }

        units = units
            .Where(unit => unit.Segments.Count > 0 && !string.IsNullOrWhiteSpace(unit.Original))
            .ToList();

        var batchSize = GetBlockTranslationConcurrency(context.Settings);
        for (var batchStart = context.ResumeUnitIndex; batchStart < units.Count; batchStart += batchSize)
        {
            var batch = units.Skip(batchStart).Take(batchSize).ToList();
            if (batch.Count == 0)
            {
                continue;
            }

            var translatedBatch = await TranslateBatchAsync(
                batch.Select(unit => new TranslationBlock(unit.Original, unit.ContextHint, unit.AdditionalRequirements)).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var batchIndex = 0; batchIndex < batch.Count; batchIndex++)
            {
                var unit = batch[batchIndex];
                var translated = translatedBatch[batchIndex];
                bilingualSegments.Add(new BilingualSegment(unit.ContextHint, unit.Original, translated));
                ApplyTranslationToUnit(unit, translated);

                var absoluteIndex = batchStart + batchIndex;
                var progress = (int)Math.Round((absoluteIndex + 1) * 100d / Math.Max(1, units.Count));
                await context.ReportProgressAsync(progress, $"{unit.ProgressLabel} {absoluteIndex + 1}/{units.Count}");
                await context.SaveCheckpointAsync(absoluteIndex + 1, 0, $"{unit.ProgressLabel} {absoluteIndex + 1}/{units.Count}");
            }
        }

        foreach (var part in partsToSave.Distinct())
        {
            part.Save();
        }

        if (context.Settings.Translation.ExportBilingualDocument)
        {
            await bilingualExportService.ExportAsync(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, bilingualSegments, context.CancellationToken);
        }
    }

    private static IEnumerable<WordTranslationUnit> BuildMainUnits(IReadOnlyList<Paragraph> selectedParagraphs)
    {
        var headingContextMap = BuildHeadingContextMap(selectedParagraphs);
        var cells = selectedParagraphs
            .Select(paragraph => paragraph.Ancestors<TableCell>().FirstOrDefault())
            .Where(cell => cell is not null)
            .Cast<TableCell>()
            .Distinct()
            .ToList();

        foreach (var cell in cells)
        {
            var unit = CreateTableCellUnit(cell, WordScope.Main, headingContextMap);
            if (unit is not null)
            {
                yield return unit;
            }
        }

        foreach (var paragraph in selectedParagraphs.Where(paragraph => !paragraph.Ancestors<TableCell>().Any()).Distinct())
        {
            var unit = CreateParagraphUnit(paragraph, WordScope.Main, headingContextMap);
            if (unit is not null)
            {
                yield return unit;
            }
        }
    }

    private static IEnumerable<WordTranslationUnit> BuildUnitsFromRootElements(
        IReadOnlyList<OpenXmlPartRootElement> roots,
        WordScope scope)
    {
        foreach (var root in roots)
        {
            foreach (var unit in BuildUnitsFromRootElement(root, scope))
            {
                yield return unit;
            }
        }
    }

    private static IEnumerable<WordTranslationUnit> BuildUnitsFromRootElement(
        OpenXmlPartRootElement root,
        WordScope scope)
    {
        var cells = root.Descendants<TableCell>().Distinct().ToList();
        foreach (var cell in cells)
        {
            var unit = CreateTableCellUnit(cell, scope, headingContextMap: null);
            if (unit is not null)
            {
                yield return unit;
            }
        }

        foreach (var paragraph in root.Descendants<Paragraph>().Where(paragraph => !paragraph.Ancestors<TableCell>().Any()).Distinct())
        {
            var unit = CreateParagraphUnit(paragraph, scope, headingContextMap: null);
            if (unit is not null)
            {
                yield return unit;
            }
        }
    }

    private static WordTranslationUnit? CreateParagraphUnit(
        Paragraph paragraph,
        WordScope scope,
        IReadOnlyDictionary<Paragraph, string>? headingContextMap)
    {
        var runs = paragraph
            .Descendants<Run>()
            .Select(run => CreateRunInfo(paragraph, run))
            .Where(info => info is not null)
            .Cast<WordRunInfo>()
            .ToList();
        if (runs.Count == 0)
        {
            return null;
        }

        var original = string.Concat(runs.Select(x => x.Original));
        if (string.IsNullOrWhiteSpace(original))
        {
            return null;
        }

        var kind = DetectParagraphKind(paragraph, scope);
        var headingContext = headingContextMap?.GetValueOrDefault(paragraph);
        return new WordTranslationUnit(
            [new WordSegmentInfo(runs, original)],
            original,
            BuildContextHint(kind, scope, headingContext),
            BuildAdditionalRequirements(kind, scope, runs),
            BuildProgressLabel(kind, scope));
    }

    private static WordTranslationUnit? CreateTableCellUnit(
        TableCell cell,
        WordScope scope,
        IReadOnlyDictionary<Paragraph, string>? headingContextMap)
    {
        var segments = cell.Elements<Paragraph>()
            .Select(paragraph =>
            {
                var runs = paragraph
                    .Descendants<Run>()
                    .Select(run => CreateRunInfo(paragraph, run))
                    .Where(info => info is not null)
                    .Cast<WordRunInfo>()
                    .ToList();
                if (runs.Count == 0)
                {
                    return null;
                }

                var original = string.Concat(runs.Select(x => x.Original));
                return string.IsNullOrWhiteSpace(original) ? null : new WordSegmentInfo(runs, original);
            })
            .Where(segment => segment is not null)
            .Cast<WordSegmentInfo>()
            .ToList();

        if (segments.Count == 0)
        {
            return null;
        }

        var original = string.Join("\n", segments.Select(x => x.Original));
        var headingContext = cell.Elements<Paragraph>()
            .Select(paragraph => headingContextMap?.GetValueOrDefault(paragraph))
            .FirstOrDefault(context => !string.IsNullOrWhiteSpace(context));
        var allRuns = segments.SelectMany(x => x.Runs).ToList();
        return new WordTranslationUnit(
            segments,
            original,
            BuildContextHint(WordContentKind.TableCell, scope, headingContext, BuildTableCellLocationHint(cell)),
            BuildAdditionalRequirements(WordContentKind.TableCell, scope, allRuns, BuildTableCellStructureHint(cell)),
            BuildProgressLabel(WordContentKind.TableCell, scope));
    }

    private static WordContentKind DetectParagraphKind(Paragraph paragraph, WordScope scope)
    {
        if (scope == WordScope.Comment)
        {
            return WordContentKind.Comment;
        }

        if (scope == WordScope.Footnote)
        {
            return WordContentKind.Footnote;
        }

        if (scope == WordScope.Endnote)
        {
            return WordContentKind.Endnote;
        }

        if (paragraph.Ancestors<TextBoxContent>().Any())
        {
            return WordContentKind.TextBox;
        }

        if (paragraph.Ancestors<TableCell>().Any())
        {
            return WordContentKind.TableCell;
        }

        if (paragraph.ParagraphProperties?.NumberingProperties is not null)
        {
            return WordContentKind.ListItem;
        }

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(styleId) &&
            styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            return WordContentKind.Heading;
        }

        if (paragraph.ParagraphProperties?.OutlineLevel is not null)
        {
            return WordContentKind.Heading;
        }

        return scope switch
        {
            WordScope.Header => WordContentKind.Header,
            WordScope.Footer => WordContentKind.Footer,
            _ => WordContentKind.Paragraph
        };
    }

    private static string BuildContextHint(
        WordContentKind kind,
        WordScope scope,
        string? headingContext,
        string? extraContext = null)
    {
        var parts = new List<string>
        {
            kind switch
            {
                WordContentKind.Heading => "Word 标题段落",
                WordContentKind.ListItem => "Word 列表项",
                WordContentKind.TableCell => "Word 表格单元格",
                WordContentKind.TextBox => "Word 文本框",
                WordContentKind.Header => "Word 页眉",
                WordContentKind.Footer => "Word 页脚",
                WordContentKind.Footnote => "Word 脚注",
                WordContentKind.Endnote => "Word 尾注",
                WordContentKind.Comment => "Word 批注",
                _ => scope == WordScope.Main ? "Word 段落" : $"Word {scope} 段落"
            }
        };

        if (!string.IsNullOrWhiteSpace(headingContext))
        {
            parts.Add($"章节上下文：{headingContext}");
        }

        if (!string.IsNullOrWhiteSpace(extraContext))
        {
            parts.Add(extraContext);
        }

        return string.Join("；", parts);
    }

    private static string BuildAdditionalRequirements(
        WordContentKind kind,
        WordScope scope,
        IReadOnlyList<WordRunInfo> runs,
        string? extraStructureHint = null)
    {
        var requirements = new List<string>();
        switch (kind)
        {
            case WordContentKind.Heading:
                requirements.Add("类型：标题。请保持标题风格，译文简洁，不要扩写成正文。");
                break;
            case WordContentKind.ListItem:
                requirements.Add("类型：列表。请保留列表层级、编号/项目符号和换行结构，不要把多个条目合并成一段。");
                break;
            case WordContentKind.TableCell:
                requirements.Add("类型：表格单元格。请保留数字、单位、百分号、缩写和换行，不要把表格短语扩写成解释性整句。");
                break;
            case WordContentKind.TextBox:
                requirements.Add("类型：文本框。请保留短句、标签和版式紧凑性，不要无故扩写。");
                break;
            case WordContentKind.Header:
            case WordContentKind.Footer:
                requirements.Add("类型：页眉页脚。请保留页码、短标题、日期和编号结构，不要扩写。");
                break;
            case WordContentKind.Footnote:
            case WordContentKind.Endnote:
                requirements.Add("类型：脚注/尾注。请保留编号、引用关系、URL、DOI、作者年份等结构。");
                break;
            case WordContentKind.Comment:
                requirements.Add("类型：批注。请保持批注语气简洁明确，不要改变批注意图。");
                break;
        }

        if (scope != WordScope.Main)
        {
            requirements.Add($"当前内容来自 Word {scope} 区域，请保持该区域常见的简短表达方式。");
        }

        if (runs.Any(run => run.HasHyperlink))
        {
            requirements.Add("当前片段包含超链接显示文本。请翻译可见文字，但保留链接文本边界，不要把链接前后的正文合并改写。");
        }

        if (runs.Any(run => run.IsSuperscript || run.IsSubscript))
        {
            requirements.Add("当前片段包含上标或下标。请保留上标/下标对应的结构边界，不要把编号、脚注标记、化学式或公式索引并入普通正文。");
        }

        if (runs.Any(run => run.IsFieldResultLike))
        {
            requirements.Add("当前片段邻近 Word 字段结果。请只翻译可见文本，不要补写字段代码或破坏引用/域结果结构。");
        }

        if (runs.Any(run => run.HasTextBoxAncestor))
        {
            requirements.Add("当前片段位于文本框中。请保持短标签、紧凑布局和独立边界，不要扩写成段落。");
        }

        if (!string.IsNullOrWhiteSpace(extraStructureHint))
        {
            requirements.Add(extraStructureHint);
        }

        return string.Join("\n", requirements);
    }

    private static string BuildProgressLabel(WordContentKind kind, WordScope scope) => kind switch
    {
        WordContentKind.TableCell => "Word 单元格",
        WordContentKind.Heading => "Word 标题",
        WordContentKind.ListItem => "Word 列表项",
        WordContentKind.TextBox => "Word 文本框",
        WordContentKind.Header => "Word 页眉",
        WordContentKind.Footer => "Word 页脚",
        WordContentKind.Footnote => "Word 脚注",
        WordContentKind.Endnote => "Word 尾注",
        WordContentKind.Comment => "Word 批注",
        _ => scope == WordScope.Main ? "Word 段落" : $"Word {scope} 段落"
    };

    private static void ApplyTranslationToUnit(WordTranslationUnit unit, string translated)
    {
        if (unit.Segments.Count == 0)
        {
            return;
        }

        if (unit.Segments.Count == 1)
        {
            ApplyTranslationToRuns(unit.Segments[0].Runs, translated);
            return;
        }

        var paragraphTranslations = TextDistributionHelper.DistributePreservingStructure(translated, unit.Segments.Count);
        for (var i = 0; i < unit.Segments.Count; i++)
        {
            ApplyTranslationToRuns(unit.Segments[i].Runs, i < paragraphTranslations.Count ? paragraphTranslations[i] : string.Empty);
        }
    }

    private static void ApplyTranslationToRuns(IReadOnlyList<WordRunInfo> runs, string translated)
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

        var formatGroups = FormattedTextRunHelper.GroupAdjacentRunsByFormat(
            runs.Select(run => new FormattedTextRun<Text>(run.Texts, run.Original, run.FormatKey)).ToList());
        if (formatGroups.Count == 1)
        {
            ApplySegmentToTexts(formatGroups[0].Texts, translated);
            return;
        }

        var segments = FormattedTextRunHelper.DistributeAcrossGroups(translated, formatGroups);
        for (var i = 0; i < formatGroups.Count; i++)
        {
            ApplySegmentToTexts(formatGroups[i].Texts, segments[i]);
        }
    }

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial)
    {
        progressService.Publish(Path.GetFileName(sourcePath), partial);
        return Task.CompletedTask;
    }

    private static IReadOnlyDictionary<Paragraph, string> BuildHeadingContextMap(IReadOnlyList<Paragraph> paragraphs)
    {
        var result = new Dictionary<Paragraph, string>();
        var headingStack = new SortedDictionary<int, string>();

        foreach (var paragraph in paragraphs)
        {
            var headingLevel = GetHeadingLevel(paragraph);
            if (headingLevel is not null)
            {
                headingStack[headingLevel.Value] = BuildParagraphPreview(paragraph);
                foreach (var key in headingStack.Keys.Where(key => key > headingLevel.Value).ToList())
                {
                    headingStack.Remove(key);
                }
            }

            if (headingStack.Count > 0)
            {
                result[paragraph] = string.Join(" > ", headingStack.OrderBy(x => x.Key).Select(x => x.Value));
            }
        }

        return result;
    }

    private static int? GetHeadingLevel(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(styleId))
        {
            var match = System.Text.RegularExpressions.Regex.Match(styleId, @"Heading\s*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out var styleLevel))
            {
                return styleLevel;
            }
        }

        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        return outlineLevel is null ? null : outlineLevel.Value + 1;
    }

    private static string BuildParagraphPreview(Paragraph paragraph)
    {
        var text = string.Concat(paragraph.Descendants<Text>().Select(textNode => textNode.Text));
        var normalized = string.Join(" ", text.Split([' ', '\r', '\n', '\t'], StringSplitOptions.RemoveEmptyEntries));
        if (normalized.Length <= 60)
        {
            return normalized;
        }

        return normalized[..60] + "...";
    }

    private static string BuildTableCellLocationHint(TableCell cell)
    {
        var row = cell.Ancestors<TableRow>().FirstOrDefault();
        var table = cell.Ancestors<Table>().FirstOrDefault();
        if (row is null || table is null)
        {
            return "该内容来自 Word 表格单元格。";
        }

        var rowIndex = table.Elements<TableRow>().TakeWhile(candidate => candidate != row).Count() + 1;
        var columnIndex = row.Elements<TableCell>().TakeWhile(candidate => candidate != cell).Count() + 1;
        return $"表格位置：第 {rowIndex} 行，第 {columnIndex} 列。";
    }

    private static string BuildTableCellStructureHint(TableCell cell)
    {
        var row = cell.Ancestors<TableRow>().FirstOrDefault();
        var table = cell.Ancestors<Table>().FirstOrDefault();
        if (row is null || table is null)
        {
            return "请保持单元格式的紧凑表达，不要把表格内容改写成整段解释。";
        }

        var rowCells = row.Elements<TableCell>().ToList();
        var rowIndex = table.Elements<TableRow>().TakeWhile(candidate => candidate != row).Count() + 1;
        var columnIndex = rowCells.TakeWhile(candidate => candidate != cell).Count() + 1;
        var rowWidth = rowCells.Count;
        return $"当前单元格位于第 {rowIndex} 行第 {columnIndex}/{Math.Max(1, rowWidth)} 列。请结合表格语境保持短标签、数值和单位的紧凑表达，不要扩写成解释性段落。";
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

    private static WordRunInfo? CreateRunInfo(Paragraph paragraph, Run run)
    {
        var texts = run.Elements<Text>().Where(x => !string.IsNullOrEmpty(x.Text)).ToList();
        if (texts.Count == 0)
        {
            return null;
        }

        return new WordRunInfo(
            texts,
            string.Concat(texts.Select(x => x.Text)),
            GetFormatKey(paragraph, run),
            HasHyperlinkAncestor(run),
            HasTextBoxAncestor(run),
            IsSuperscript(run),
            IsSubscript(run),
            IsFieldResultLike(run));
    }

    private static string GetFormatKey(Paragraph paragraph, Run run)
    {
        var runProperties = run.RunProperties?.OuterXml ?? string.Empty;
        var paragraphProperties = paragraph.ParagraphProperties?.OuterXml ?? string.Empty;
        return $"{paragraphProperties}|{runProperties}|{GetRunBoundaryKey(run)}";
    }

    private static string GetRunBoundaryKey(Run run)
    {
        var flags = new List<string>(4);
        if (HasHyperlinkAncestor(run))
        {
            flags.Add("link");
        }

        if (HasTextBoxAncestor(run))
        {
            flags.Add("textbox");
        }

        if (IsSuperscript(run))
        {
            flags.Add("sup");
        }
        else if (IsSubscript(run))
        {
            flags.Add("sub");
        }

        if (IsFieldResultLike(run))
        {
            flags.Add("field");
        }

        return flags.Count == 0 ? "plain" : string.Join("|", flags);
    }

    private static bool HasHyperlinkAncestor(OpenXmlElement element) =>
        element.Ancestors<Hyperlink>().Any();

    private static bool HasTextBoxAncestor(OpenXmlElement element) =>
        element.Ancestors<TextBoxContent>().Any();

    private static bool IsSuperscript(Run run) =>
        run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript;

    private static bool IsSubscript(Run run) =>
        run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript;

    private static bool IsFieldResultLike(Run run) =>
        run.Elements<FieldChar>().Any() ||
        run.Elements<FieldCode>().Any() ||
        run.Ancestors<SimpleField>().Any();

    private static void ApplySegmentToTexts(IReadOnlyList<Text> texts, string segment)
    {
        if (texts.Count == 0)
        {
            return;
        }

        var originalFirstText = texts[0].Text ?? string.Empty;
        var leadingSpace = WhitespacePreservationHelper.GetLeadingWhitespace(originalFirstText);
        var trailingSpace = WhitespacePreservationHelper.GetTrailingWhitespace(texts[^1].Text ?? string.Empty);

        var processedSegment = leadingSpace + segment.Trim() + trailingSpace;

        texts[0].Space = SpaceProcessingModeValues.Preserve;
        texts[0].Text = processedSegment;

        for (var i = 1; i < texts.Count; i++)
        {
            texts[i].Space = SpaceProcessingModeValues.Preserve;
            texts[i].Text = string.Empty;
        }
    }

    private enum WordScope
    {
        Main,
        Header,
        Footer,
        Footnote,
        Endnote,
        Comment
    }

    private enum WordContentKind
    {
        Paragraph,
        Heading,
        ListItem,
        TableCell,
        TextBox,
        Header,
        Footer,
        Footnote,
        Endnote,
        Comment
    }

    private sealed record WordRunInfo(
        IReadOnlyList<Text> Texts,
        string Original,
        string FormatKey,
        bool HasHyperlink,
        bool HasTextBoxAncestor,
        bool IsSuperscript,
        bool IsSubscript,
        bool IsFieldResultLike);
    private sealed record WordSegmentInfo(IReadOnlyList<WordRunInfo> Runs, string Original);
    private sealed record WordTranslationUnit(
        IReadOnlyList<WordSegmentInfo> Segments,
        string Original,
        string ContextHint,
        string AdditionalRequirements,
        string ProgressLabel);
    private sealed record WordPageInfo(Paragraph Paragraph, int PageNumber);
}
