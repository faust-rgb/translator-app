using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public sealed class EbookDocumentTranslator(
    ITextTranslationService textTranslationService,
    IAppLogService logService,
    ITranslationProgressService progressService,
    IBilingualExportService bilingualExportService,
    IEbookConversionService ebookConversionService,
    IEbookDocxExportService ebookDocxExportService)
    : DocumentTranslatorBase(textTranslationService, logService)
{
    private const int LongParagraphSplitThreshold = 420;
    private static readonly Regex SentenceSplitRegex = new(@"(?<=[\.!\?。！？；;:：])(?=\s|$)|(?<=\n)(?=\S)", RegexOptions.Compiled);
    private static readonly HashSet<string> BlockElementNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "title",
        "p",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "li",
        "blockquote",
        "dd",
        "dt",
        "figcaption",
        "caption",
        "td",
        "th",
        "div"
    };

    private static readonly HashSet<string> ExcludedAncestorNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "script",
        "style",
        "code",
        "pre",
        "math",
        "svg",
        "rt",
        "rp"
    };

    public override bool CanHandle(string extension) =>
        extension is ".epub" or ".mobi" or ".azw3";

    public override async Task TranslateAsync(TranslationJobContext context)
    {
        ArgumentNullException.ThrowIfNull(context);

        var targetExtension = ResolveTargetExtension(context.Settings.Translation.EbookOutputFormat);
        var outputPath = BuildOutputPath(context.Item.SourcePath, context.Settings.Translation.OutputDirectory, targetExtension);
        context.Item.OutputPath = outputPath;
        var workingDirectory = GetWorkingDirectory(context.Item.SourcePath);
        Directory.CreateDirectory(workingDirectory);
        var completed = false;

        try
        {
            var sourceEpubPath = await NormalizeSourceToEpubAsync(context, workingDirectory);
            var bilingualSegments = new List<BilingualSegment>();
            var workspace = await TranslateIntoWorkspaceAsync(sourceEpubPath, workingDirectory, context, bilingualSegments);

            if (string.Equals(targetExtension, ".epub", StringComparison.OrdinalIgnoreCase))
            {
                PackDirectoryAsEpub(workspace.ExtractDirectory, outputPath);
            }
            else
            {
                await ebookDocxExportService.ExportAsync(
                    outputPath,
                    workspace.BookTitle,
                    workspace.Cover,
                    workspace.Metadata,
                    workspace.ContentDocuments,
                    context.CancellationToken);
            }

            if (context.Settings.Translation.ExportBilingualDocument)
            {
                await bilingualExportService.ExportAsync(
                    context.Item.SourcePath,
                    context.Settings.Translation.OutputDirectory,
                    bilingualSegments,
                    context.CancellationToken);
            }

            completed = true;
        }
        finally
        {
            if (completed)
            {
                TryDeleteDirectory(workingDirectory);
            }
            else if (Directory.Exists(workingDirectory))
            {
                Log($"电子书工作区已保留，便于下次继续：{workingDirectory}");
            }
        }
    }

    private async Task<string> NormalizeSourceToEpubAsync(TranslationJobContext context, string workingDirectory)
    {
        var extension = Path.GetExtension(context.Item.SourcePath).ToLowerInvariant();
        if (extension == ".epub")
        {
            return context.Item.SourcePath;
        }

        Log($"{extension} 电子书需要先转换为 EPUB 才能继续翻译。");
        var normalizedEpubPath = Path.Combine(workingDirectory, "source.epub");
        if (context.ResumeUnitIndex > 0 && File.Exists(normalizedEpubPath))
        {
            Log("检测到上次保留的已转换 EPUB，继续复用。");
            return normalizedEpubPath;
        }

        await ebookConversionService.ConvertAsync(
            context.Item.SourcePath,
            normalizedEpubPath,
            context.Settings.Translation.CalibreExecutablePath,
            context.CancellationToken);
        return normalizedEpubPath;
    }

    private async Task<EpubWorkspace> TranslateIntoWorkspaceAsync(
        string sourceEpubPath,
        string workingDirectory,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments)
    {
        var extractDirectory = Path.Combine(workingDirectory, "epub-src");
        if (context.ResumeUnitIndex > 0 && Directory.Exists(extractDirectory))
        {
            Log($"从已保留的 EPUB 工作区继续：已完成 {context.ResumeUnitIndex} 个单元。");
        }
        else
        {
            TryDeleteDirectory(extractDirectory);
            ZipFile.ExtractToDirectory(sourceEpubPath, extractDirectory);
        }

        var packagePath = ResolvePackagePath(extractDirectory);
        var packageDocument = XDocument.Load(packagePath, LoadOptions.PreserveWhitespace);
        var bookTitle = ResolveBookTitle(packageDocument, context.Item.SourcePath);
        var metadata = ResolveMetadata(packageDocument, bookTitle);
        var contentFiles = ResolveContentDocumentPaths(packagePath, packageDocument);
        var cover = ResolveCoverInfo(packagePath, packageDocument, contentFiles);
        var navFiles = ResolveNavigationDocumentPaths(packagePath, packageDocument);
        var ncxFiles = ResolveNcxDocumentPaths(packagePath, packageDocument);
        var requestedRange = GetRequestedRange(context.Settings, contentFiles.Count);
        var fullContentRangeSelected = requestedRange.Start == 1 && requestedRange.End >= contentFiles.Count;
        var selectedContentFiles = contentFiles
            .Select((path, index) => new { Path = path, Index = index + 1 })
            .Where(x => IsWithinRequestedRange(x.Index, requestedRange))
            .ToList();

        var contentDocuments = selectedContentFiles
            .Select(x => LoadContentDocument(x.Path, x.Index, contentFiles.Count))
            .ToList();
        var navigationDocuments = navFiles
            .Where(path => contentFiles.All(content => !PathsEqual(content, path)))
            .Select(LoadContentDocument)
            .Where(static x => x.Units.Count > 0)
            .ToList();
        var ncxDocuments = ncxFiles
            .Select(path => new NcxDocument(path, GetNcxUnitCount(path)))
            .Where(static x => x.UnitCount > 0)
            .ToList();

        var totalUnits = Math.Max(
            1,
            contentDocuments.Sum(x => x.Units.Count) +
            navigationDocuments.Sum(x => x.Units.Count) +
            ncxDocuments.Sum(x => x.UnitCount));
        var resumeUnitIndex = Math.Clamp(context.ResumeUnitIndex, 0, totalUnits);
        var processedUnits = 0;

        if (resumeUnitIndex > 0)
        {
            var resumeProgress = (int)Math.Round(resumeUnitIndex * 100d / totalUnits);
            await context.ReportProgressAsync(resumeProgress, $"电子书恢复进度 {resumeUnitIndex}/{totalUnits}");
        }

        foreach (var document in contentDocuments)
        {
            processedUnits = await TranslateContentDocumentAsync(document, context, bilingualSegments, totalUnits, processedUnits, resumeUnitIndex);
            document.Document.Save(document.Path, SaveOptions.DisableFormatting);
        }

        var headingMap = BuildHeadingTargetMap(contentDocuments);

        foreach (var navigationDocument in navigationDocuments)
        {
            processedUnits = await TranslateNavigationDocumentAsync(navigationDocument, headingMap, fullContentRangeSelected, context, bilingualSegments, totalUnits, processedUnits, resumeUnitIndex);
        }

        foreach (var ncxDocument in ncxDocuments)
        {
            processedUnits = await TranslateNcxDocumentAsync(ncxDocument.Path, headingMap, fullContentRangeSelected, context, bilingualSegments, totalUnits, processedUnits, resumeUnitIndex);
        }

        var translatedContentDocuments = selectedContentFiles
            .Select(x => new EpubExportDocument(x.Path, XDocument.Load(x.Path, LoadOptions.PreserveWhitespace)))
            .ToList();

        return new EpubWorkspace(extractDirectory, bookTitle, cover, metadata, translatedContentDocuments);
    }

    private async Task<int> TranslateContentDocumentAsync(
        EpubContentDocument document,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments,
        int totalUnits,
        int processedUnits,
        int resumeUnitIndex)
    {
        var documentResumeOffset = Math.Min(document.Units.Count, Math.Max(0, resumeUnitIndex - processedUnits));
        var batchSize = GetBlockTranslationConcurrency(context.Settings);
        for (var batchStart = documentResumeOffset; batchStart < document.Units.Count; batchStart += batchSize)
        {
            var batch = document.Units.Skip(batchStart).Take(batchSize).ToList();
            var fragmentBlocks = new List<TranslationBlock>();
            var fragmentOffsets = new List<(int Start, int Count)>();
            foreach (var unit in batch)
            {
                var start = fragmentBlocks.Count;
                foreach (var fragment in unit.Fragments)
                {
                    fragmentBlocks.Add(new TranslationBlock(fragment.Original, fragment.ContextHint, fragment.AdditionalRequirements));
                }

                fragmentOffsets.Add((start, unit.Fragments.Count));
            }

            var translatedFragments = await TranslateBatchAsync(
                fragmentBlocks,
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var index = 0; index < batch.Count; index++)
            {
                var unit = batch[index];
                var translated = string.Concat(
                    translatedFragments
                        .Skip(fragmentOffsets[index].Start)
                        .Take(fragmentOffsets[index].Count));
                ApplyTranslatedText(unit.TextNodes, translated);
                bilingualSegments.Add(new BilingualSegment(unit.ContextHint, unit.Original, translated));

                var absoluteUnitIndex = processedUnits + batchStart + index + 1;
                var progress = (int)Math.Round(absoluteUnitIndex * 100d / totalUnits);
                await context.ReportProgressAsync(progress, $"电子书 {Path.GetFileName(document.Path)} {absoluteUnitIndex}/{totalUnits}");
                await context.SaveCheckpointAsync(absoluteUnitIndex, 0, $"电子书 {Path.GetFileName(document.Path)} {absoluteUnitIndex}/{totalUnits}");
            }
        }

        return processedUnits + document.Units.Count;
    }

    private async Task<int> TranslateNavigationDocumentAsync(
        EpubContentDocument document,
        IReadOnlyDictionary<string, string> headingMap,
        bool allowFallbackTranslation,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments,
        int totalUnits,
        int processedUnits,
        int resumeUnitIndex)
    {
        var pendingUnits = new List<(EpubTranslationUnit Unit, int AbsoluteIndex, string? SynchronizedTitle)>();
        var fallbackUnits = new List<(EpubTranslationUnit Unit, int AbsoluteIndex)>();
        for (var index = 0; index < document.Units.Count; index++)
        {
            var absoluteIndex = processedUnits + index + 1;
            if (absoluteIndex <= resumeUnitIndex)
            {
                continue;
            }

            var unit = document.Units[index];
            var href = FindOwningHref(unit.TextNodes);
            var targetKey = string.IsNullOrWhiteSpace(href) ? null : ResolveNavigationTargetKey(document.Path, href!);
            if (targetKey is not null && TryResolveHeadingTitle(headingMap, targetKey, out var synchronizedTitle))
            {
                pendingUnits.Add((unit, absoluteIndex, synchronizedTitle));
            }
            else if (allowFallbackTranslation)
            {
                fallbackUnits.Add((unit, absoluteIndex));
            }
        }

        var fallbackTranslations = new Dictionary<int, string>();
        if (fallbackUnits.Count > 0)
        {
            var translatedBatch = await TranslateBatchAsync(
                fallbackUnits.Select(x => new TranslationBlock(x.Unit.Original, x.Unit.ContextHint, x.Unit.AdditionalRequirements)).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var index = 0; index < fallbackUnits.Count; index++)
            {
                fallbackTranslations[fallbackUnits[index].AbsoluteIndex] = translatedBatch[index];
            }
        }

        var synchronizedTranslations = pendingUnits.ToDictionary(x => x.AbsoluteIndex, x => x.SynchronizedTitle!);
        foreach (var unit in document.Units.Select((item, index) => new { Unit = item, AbsoluteIndex = processedUnits + index + 1 }))
        {
            if (unit.AbsoluteIndex <= resumeUnitIndex)
            {
                continue;
            }

            if (!synchronizedTranslations.TryGetValue(unit.AbsoluteIndex, out var translated) &&
                !fallbackTranslations.TryGetValue(unit.AbsoluteIndex, out translated))
            {
                continue;
            }

            ApplyTranslatedText(unit.Unit.TextNodes, translated);
            bilingualSegments.Add(new BilingualSegment(unit.Unit.ContextHint, unit.Unit.Original, translated));
            var progress = (int)Math.Round(unit.AbsoluteIndex * 100d / totalUnits);
            await context.ReportProgressAsync(progress, $"电子书目录 {unit.AbsoluteIndex}/{totalUnits}");
            await context.SaveCheckpointAsync(unit.AbsoluteIndex, 0, $"电子书目录 {unit.AbsoluteIndex}/{totalUnits}");
        }

        document.Document.Save(document.Path, SaveOptions.DisableFormatting);
        return processedUnits + document.Units.Count;
    }

    private async Task<int> TranslateNcxDocumentAsync(
        string ncxPath,
        IReadOnlyDictionary<string, string> headingMap,
        bool allowFallbackTranslation,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments,
        int totalUnits,
        int processedUnits,
        int resumeUnitIndex)
    {
        var document = XDocument.Load(ncxPath, LoadOptions.PreserveWhitespace);
        var textElements = document
            .Descendants()
            .Where(x => x.Name.LocalName == "text")
            .Where(x => !string.IsNullOrWhiteSpace(x.Value))
            .ToList();

        if (textElements.Count == 0)
        {
            return 0;
        }

        var synchronizedElements = new List<(XElement Element, int AbsoluteIndex, string Translation)>();
        var fallbackElements = new List<(XElement Element, int Index, int AbsoluteIndex)>();
        for (var index = 0; index < textElements.Count; index++)
        {
            var absoluteIndex = processedUnits + index + 1;
            if (absoluteIndex <= resumeUnitIndex)
            {
                continue;
            }

            var element = textElements[index];
            var navPoint = element.Ancestors().FirstOrDefault(x => x.Name.LocalName == "navPoint");
            var src = navPoint?
                .Descendants()
                .FirstOrDefault(x => x.Name.LocalName == "content")
                ?.Attribute("src")
                ?.Value;
            var targetKey = string.IsNullOrWhiteSpace(src) ? null : ResolveNavigationTargetKey(ncxPath, src!);
            if (targetKey is not null && TryResolveHeadingTitle(headingMap, targetKey, out var synchronizedTitle))
            {
                synchronizedElements.Add((element, absoluteIndex, synchronizedTitle));
            }
            else if (allowFallbackTranslation)
            {
                fallbackElements.Add((element, index, absoluteIndex));
            }
        }

        var fallbackTranslations = new Dictionary<int, string>();
        if (fallbackElements.Count > 0)
        {
            var translatedBatch = await TranslateBatchAsync(
                fallbackElements.Select(x => new TranslationBlock(x.Element.Value, $"EPUB 目录 {x.Index + 1}", "类型：电子书导航目录。请保持简洁，尽量与对应正文标题风格一致。")).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var index = 0; index < fallbackElements.Count; index++)
            {
                fallbackTranslations[fallbackElements[index].AbsoluteIndex] = translatedBatch[index];
            }
        }

        var synchronizedTranslations = synchronizedElements.ToDictionary(x => x.AbsoluteIndex, x => x.Translation);
        foreach (var item in textElements.Select((element, index) => new { Element = element, AbsoluteIndex = processedUnits + index + 1, DisplayIndex = index + 1 }))
        {
            if (item.AbsoluteIndex <= resumeUnitIndex)
            {
                continue;
            }

            if (!synchronizedTranslations.TryGetValue(item.AbsoluteIndex, out var translated) &&
                !fallbackTranslations.TryGetValue(item.AbsoluteIndex, out translated))
            {
                continue;
            }

            bilingualSegments.Add(new BilingualSegment($"EPUB 目录 {item.DisplayIndex}", item.Element.Value, translated));
            item.Element.Value = translated;
            var progress = (int)Math.Round(item.AbsoluteIndex * 100d / totalUnits);
            await context.ReportProgressAsync(progress, $"电子书导航 {item.AbsoluteIndex}/{totalUnits}");
            await context.SaveCheckpointAsync(item.AbsoluteIndex, 0, $"电子书导航 {item.AbsoluteIndex}/{totalUnits}");
        }

        document.Save(ncxPath, SaveOptions.DisableFormatting);
        return processedUnits + textElements.Count;
    }

    private static Dictionary<string, string> BuildHeadingTargetMap(IReadOnlyList<EpubContentDocument> contentDocuments)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var document in contentDocuments)
        {
            var headings = document.Document
                .Descendants()
                .Where(IsPotentialHeadingElement)
                .ToList();
            var firstHeading = headings
                .Select(GetElementText)
                .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));
            if (string.IsNullOrWhiteSpace(firstHeading))
            {
                firstHeading = document.DocumentTitle;
            }

            if (!string.IsNullOrWhiteSpace(firstHeading))
            {
                map[NormalizeDocumentKey(document.Path)] = firstHeading!;
            }

            foreach (var heading in headings)
            {
                var id = heading.Attribute("id")?.Value ??
                         heading.Attribute(XName.Get("id", "http://www.w3.org/XML/1998/namespace"))?.Value;
                var text = GetElementText(heading);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(id))
                {
                    map[$"{NormalizeDocumentKey(document.Path)}#{NormalizeAnchorFragment(id)}"] = text;
                }
            }
        }

        return map;
    }

    private static bool IsHeadingElement(XElement element) =>
        element.Name.LocalName is "h1" or "h2" or "h3" or "h4" or "h5" or "h6";

    private static bool IsPotentialHeadingElement(XElement element)
    {
        if (IsHeadingElement(element) || string.Equals(element.Name.LocalName, "title", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var hint = string.Join(
            " ",
            element.Attribute("class")?.Value ?? string.Empty,
            element.Attribute("epub:type")?.Value ?? string.Empty,
            element.Attribute("type")?.Value ?? string.Empty,
            element.Attribute("id")?.Value ?? string.Empty);

        return hint.Contains("chapter", StringComparison.OrdinalIgnoreCase) ||
               hint.Contains("title", StringComparison.OrdinalIgnoreCase) ||
               hint.Contains("heading", StringComparison.OrdinalIgnoreCase) ||
               hint.Contains("subtitle", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetElementText(XElement element) =>
        string.Concat(element
            .DescendantNodes()
            .OfType<XText>()
            .Select(x => x.Value))
            .Trim();

    private static string? FindOwningHref(IReadOnlyList<XText> textNodes)
    {
        foreach (var node in textNodes)
        {
            var link = node.Ancestors().FirstOrDefault(x => x.Name.LocalName == "a");
            var href = link?.Attribute("href")?.Value;
            if (!string.IsNullOrWhiteSpace(href))
            {
                return href;
            }
        }

        return null;
    }

    private static string? ResolveNavigationTargetKey(string navigationDocumentPath, string href)
    {
        if (string.IsNullOrWhiteSpace(href) || href.StartsWith("http", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var normalizedHref = href.Trim();
        var queryIndex = normalizedHref.IndexOf('?');
        if (queryIndex >= 0)
        {
            normalizedHref = normalizedHref[..queryIndex];
        }

        var hashIndex = normalizedHref.IndexOf('#');
        var filePart = hashIndex >= 0 ? normalizedHref[..hashIndex] : normalizedHref;
        var fragment = hashIndex >= 0 ? normalizedHref[(hashIndex + 1)..] : string.Empty;
        var baseDirectory = Path.GetDirectoryName(navigationDocumentPath)!;
        var targetPath = string.IsNullOrWhiteSpace(filePart)
            ? NormalizeDocumentKey(navigationDocumentPath)
            : NormalizePackageRelativePath(baseDirectory, filePart);
        var normalizedFragment = NormalizeAnchorFragment(fragment);

        return string.IsNullOrWhiteSpace(normalizedFragment)
            ? targetPath
            : $"{targetPath}#{normalizedFragment}";
    }

    private static bool TryResolveHeadingTitle(IReadOnlyDictionary<string, string> headingMap, string targetKey, out string synchronizedTitle)
    {
        if (headingMap.TryGetValue(targetKey, out synchronizedTitle!))
        {
            return true;
        }

        var hashIndex = targetKey.IndexOf('#');
        if (hashIndex > 0)
        {
            var documentKey = targetKey[..hashIndex];
            if (headingMap.TryGetValue(documentKey, out synchronizedTitle!))
            {
                return true;
            }
        }

        synchronizedTitle = string.Empty;
        return false;
    }

    private static EpubCoverInfo? ResolveCoverInfo(
        string packagePath,
        XDocument packageDocument,
        IReadOnlyList<string> contentFiles)
    {
        var manifestItems = packageDocument
            .Descendants()
            .Where(x => x.Name.LocalName == "item")
            .Select(item => new ManifestItem(
                item.Attribute("id")?.Value ?? string.Empty,
                item.Attribute("href")?.Value ?? string.Empty,
                item.Attribute("media-type")?.Value ?? string.Empty,
                item.Attribute("properties")?.Value ?? string.Empty))
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.Href))
            .ToList();
        var packageDirectory = Path.GetDirectoryName(packagePath)!;

        var coverImageItem = manifestItems.FirstOrDefault(x => x.Properties.Contains("cover-image", StringComparison.OrdinalIgnoreCase));
        if (coverImageItem is null)
        {
            var legacyCoverId = packageDocument
                .Descendants()
                .FirstOrDefault(x => x.Name.LocalName == "meta" &&
                                     string.Equals(x.Attribute("name")?.Value, "cover", StringComparison.OrdinalIgnoreCase))
                ?.Attribute("content")
                ?.Value;
            if (!string.IsNullOrWhiteSpace(legacyCoverId))
            {
                coverImageItem = manifestItems.FirstOrDefault(x => string.Equals(x.Id, legacyCoverId, StringComparison.Ordinal));
            }
        }

        string? coverImagePath = null;
        if (coverImageItem is not null)
        {
            coverImagePath = NormalizePackageRelativePath(packageDirectory, coverImageItem.Href);
        }

        var coverDocumentPath = contentFiles.FirstOrDefault(path =>
            path.Contains("cover", StringComparison.OrdinalIgnoreCase) ||
            (!string.IsNullOrWhiteSpace(coverImagePath) && DocumentReferencesImage(path, coverImagePath!)));

        return coverImagePath is null && coverDocumentPath is null
            ? null
            : new EpubCoverInfo(coverImagePath, coverDocumentPath);
    }

    private static bool DocumentReferencesImage(string documentPath, string imagePath)
    {
        if (!File.Exists(documentPath))
        {
            return false;
        }

        var document = XDocument.Load(documentPath, LoadOptions.PreserveWhitespace);
        return document
            .Descendants()
            .Where(x => x.Name.LocalName == "img")
            .Select(x => x.Attribute("src")?.Value)
            .Where(src => !string.IsNullOrWhiteSpace(src))
            .Any(src => string.Equals(
                NormalizePackageRelativePath(Path.GetDirectoryName(documentPath)!, src!),
                Path.GetFullPath(imagePath),
                StringComparison.OrdinalIgnoreCase));
    }

    private sealed record ManifestItem(string Id, string Href, string MediaType, string Properties);

    private static EpubContentDocument LoadContentDocument(string path)
    {
        var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        var documentTitle = ResolveDocumentTitle(document);
        var units = BuildTranslationUnits(document, Path.GetFileName(path), documentTitle, null);
        return new EpubContentDocument(path, document, units, documentTitle);
    }

    private static EpubContentDocument LoadContentDocument(string path, int contentIndex, int totalContentCount)
    {
        var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        var documentTitle = ResolveDocumentTitle(document);
        var units = BuildTranslationUnits(
            document,
            Path.GetFileName(path),
            documentTitle,
            $"章节 {contentIndex}/{totalContentCount}");
        return new EpubContentDocument(path, document, units, documentTitle);
    }

    private static int GetNcxUnitCount(string path)
    {
        var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        return document
            .Descendants()
            .Count(x => x.Name.LocalName == "text" && !string.IsNullOrWhiteSpace(x.Value));
    }

    private static List<EpubTranslationUnit> BuildTranslationUnits(
        XDocument document,
        string fileName,
        string documentTitle,
        string? chapterIndexLabel)
    {
        var units = new List<EpubTranslationUnit>();
        var candidates = document
            .Descendants()
            .Where(IsTranslatableLeafElement)
            .ToList();
        var headingCandidates = document
            .Descendants()
            .Where(IsPotentialHeadingElement)
            .ToList();

        foreach (var element in candidates)
        {
            var textNodes = element
                .DescendantNodes()
                .OfType<XText>()
                .Where(node => !string.IsNullOrWhiteSpace(node.Value))
                .Where(node => node.Ancestors().All(ancestor => !ExcludedAncestorNames.Contains(ancestor.Name.LocalName)))
                .ToList();

            if (textNodes.Count == 0)
            {
                continue;
            }

            var original = string.Concat(textNodes.Select(x => x.Value));
            if (!LooksTranslatable(original))
            {
                continue;
            }

            var headingContext = ResolveHeadingContext(element, headingCandidates, documentTitle);
            var contextHint = BuildContextHint(fileName, element, headingContext, chapterIndexLabel);
            var additionalRequirements = BuildAdditionalRequirements(element);
            var fragments = BuildTranslationFragments(original, contextHint, additionalRequirements);
            units.Add(new EpubTranslationUnit(textNodes, original, contextHint, additionalRequirements, fragments));
        }

        return units;
    }

    private static bool IsTranslatableLeafElement(XElement element)
    {
        if (!BlockElementNames.Contains(element.Name.LocalName))
        {
            return false;
        }

        if (element.Ancestors().Any(ancestor => ExcludedAncestorNames.Contains(ancestor.Name.LocalName)))
        {
            return false;
        }

        return !element.Descendants().Any(descendant => descendant != element && BlockElementNames.Contains(descendant.Name.LocalName));
    }

    private static bool LooksTranslatable(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value.Trim().Any(char.IsLetter);
    }

    private static string ResolveDocumentTitle(XDocument document)
    {
        var title = document
            .Descendants()
            .FirstOrDefault(x => string.Equals(x.Name.LocalName, "title", StringComparison.OrdinalIgnoreCase))
            ?.Value
            ?.Trim();

        if (!string.IsNullOrWhiteSpace(title))
        {
            return title;
        }

        var firstHeading = document
            .Descendants()
            .Where(IsPotentialHeadingElement)
            .Select(GetElementText)
            .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));

        return firstHeading ?? string.Empty;
    }

    private static string ResolveHeadingContext(XElement element, IReadOnlyList<XElement> headingCandidates, string documentTitle)
    {
        var orderedElements = element.Document?.Descendants().ToList() ?? [];
        if (orderedElements.Count == 0)
        {
            return documentTitle;
        }

        var positions = orderedElements
            .Select((item, index) => new { item, index })
            .ToDictionary(x => x.item, x => x.index);
        if (!positions.TryGetValue(element, out var targetIndex))
        {
            return documentTitle;
        }

        var resolvedHeading = documentTitle;
        foreach (var candidate in headingCandidates)
        {
            if (!positions.TryGetValue(candidate, out var headingIndex) || headingIndex > targetIndex)
            {
                continue;
            }

            var text = GetElementText(candidate);
            if (!string.IsNullOrWhiteSpace(text))
            {
                resolvedHeading = text;
            }
        }

        return resolvedHeading;
    }

    private static string BuildContextHint(string fileName, XElement element, string headingContext, string? chapterIndexLabel)
    {
        var parts = new List<string> { $"EPUB {DescribeElementRole(element)} <{element.Name.LocalName}>" };
        if (!string.IsNullOrWhiteSpace(chapterIndexLabel))
        {
            parts.Add(chapterIndexLabel);
        }

        if (!string.IsNullOrWhiteSpace(headingContext))
        {
            parts.Add($"章节上下文：{headingContext}");
        }

        parts.Add($"来源文件：{fileName}");
        return string.Join("；", parts);
    }

    private static string DescribeElementRole(XElement element) => element.Name.LocalName switch
    {
        "title" => "文档标题",
        "h1" or "h2" or "h3" or "h4" or "h5" or "h6" => "章节标题",
        "li" => "列表项",
        "figcaption" or "caption" => "图表说明",
        "td" or "th" => "表格单元格",
        "blockquote" => "引用块",
        _ => "正文块"
    };

    private static string BuildAdditionalRequirements(XElement element)
    {
        var requirements = new List<string>();
        switch (element.Name.LocalName)
        {
            case "title":
            case "h1":
            case "h2":
            case "h3":
            case "h4":
            case "h5":
            case "h6":
                requirements.Add("类型：电子书标题。请保持标题风格和层级感，译文简洁，不要扩写成正文。");
                break;
            case "li":
            case "dd":
            case "dt":
                requirements.Add("类型：电子书列表。请保留条目边界、编号/项目符号和换行结构，不要把多个条目合并成一段。");
                break;
            case "figcaption":
            case "caption":
                requirements.Add("类型：电子书图表说明。请保留图号、子图编号、单位、括号和短说明风格。");
                break;
            case "td":
            case "th":
                requirements.Add("类型：电子书表格单元格。请保留数字、单位、缩写和换行，不要把短语扩写成解释性整句。");
                break;
            case "blockquote":
                requirements.Add("类型：引用块。请保留引用语气和段落边界，不要混入解释性说明。");
                break;
            default:
                requirements.Add("类型：电子书正文。请保留段落边界、内联强调边界和链接文本边界，不要无故扩写。");
                break;
        }

        if (element.Descendants().Any(x => x.Name.LocalName == "a"))
        {
            requirements.Add("当前片段包含链接文本。请翻译可见文字，但不要破坏链接边界，也不要把链接前后的内容并到一起。");
        }

        if (element.Descendants().Any(x => x.Name.LocalName is "em" or "strong" or "b" or "i"))
        {
            requirements.Add("当前片段包含强调样式。请尽量保持强调部分的语义边界，不要把强调片段并入相邻短语。");
        }

        return string.Join("\n", requirements);
    }

    private static IReadOnlyList<EpubTranslationFragment> BuildTranslationFragments(
        string original,
        string contextHint,
        string additionalRequirements)
    {
        var fragments = SplitLongTextForTranslation(original)
            .Select(fragment => new EpubTranslationFragment(fragment, contextHint, additionalRequirements))
            .ToList();

        if (fragments.Count == 0)
        {
            fragments.Add(new EpubTranslationFragment(original, contextHint, additionalRequirements));
        }

        return fragments;
    }

    private static IReadOnlyList<string> SplitLongTextForTranslation(string original)
    {
        if (string.IsNullOrWhiteSpace(original) || original.Length <= LongParagraphSplitThreshold)
        {
            return new[] { original };
        }

        var rawSegments = SentenceSplitRegex
            .Split(original)
            .Select(segment => segment)
            .Where(segment => !string.IsNullOrWhiteSpace(segment))
            .ToList();

        if (rawSegments.Count <= 1)
        {
            return new[] { original };
        }

        var merged = new List<string>();
        var builder = new StringBuilder();
        foreach (var segment in rawSegments)
        {
            if (builder.Length > 0 && builder.Length + segment.Length > LongParagraphSplitThreshold)
            {
                merged.Add(builder.ToString());
                builder.Clear();
            }

            builder.Append(segment);
        }

        if (builder.Length > 0)
        {
            merged.Add(builder.ToString());
        }

        return merged.Count == 0 ? new[] { original } : merged;
    }

    private static void ApplyTranslatedText(IReadOnlyList<XText> textNodes, string translated)
    {
        if (textNodes.Count == 0)
        {
            return;
        }

        if (textNodes.Count == 1)
        {
            textNodes[0].Value = WhitespacePreservationHelper.PreserveEdgeWhitespace(textNodes[0].Value, translated);
            return;
        }

        var segments = TextDistributionHelper.Distribute(
            translated,
            textNodes.Select(node => Math.Max(1, node.Value.Trim().Length)).ToList());

        for (var index = 0; index < textNodes.Count; index++)
        {
            textNodes[index].Value = WhitespacePreservationHelper.PreserveEdgeWhitespace(textNodes[index].Value, segments[index]);
        }
    }

    private static string ResolvePackagePath(string extractDirectory)
    {
        var containerPath = Path.Combine(extractDirectory, "META-INF", "container.xml");
        var container = XDocument.Load(containerPath, LoadOptions.PreserveWhitespace);
        var rootFile = container
            .Descendants()
            .FirstOrDefault(x => x.Name.LocalName == "rootfile")
            ?.Attribute("full-path")
            ?.Value;

        if (string.IsNullOrWhiteSpace(rootFile))
        {
            throw new InvalidOperationException("EPUB 缺少 OPF 包文档，无法继续处理。");
        }

        return NormalizePackageRelativePath(extractDirectory, rootFile);
    }

    private static string ResolveBookTitle(XDocument packageDocument, string sourcePath)
    {
        var title = packageDocument
            .Descendants()
            .FirstOrDefault(x => x.Name.LocalName == "title")
            ?.Value
            ?.Trim();

        return string.IsNullOrWhiteSpace(title)
            ? Path.GetFileNameWithoutExtension(sourcePath)
            : title;
    }

    private static EpubMetadata ResolveMetadata(XDocument packageDocument, string fallbackTitle)
    {
        var metadataElement = packageDocument.Descendants().FirstOrDefault(x => x.Name.LocalName == "metadata");

        string FirstValue(params string[] names) =>
            metadataElement?
                .Elements()
                .FirstOrDefault(x => names.Contains(x.Name.LocalName, StringComparer.OrdinalIgnoreCase))
                ?.Value
                ?.Trim() ?? string.Empty;

        var creators = metadataElement?
            .Elements()
            .Where(x => string.Equals(x.Name.LocalName, "creator", StringComparison.OrdinalIgnoreCase))
            .Select(x => x.Value?.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Cast<string>()
            .ToList() ?? [];

        return new EpubMetadata(
            string.IsNullOrWhiteSpace(FirstValue("title")) ? fallbackTitle : FirstValue("title"),
            creators,
            FirstValue("publisher"),
            FirstValue("language"),
            FirstValue("date", "modified"),
            FirstValue("identifier"),
            FirstValue("description"));
    }

    private static IReadOnlyList<string> ResolveContentDocumentPaths(string packagePath, XDocument packageDocument)
    {
        var manifestById = packageDocument
            .Descendants()
            .Where(x => x.Name.LocalName == "item")
            .Select(item => new
            {
                Id = item.Attribute("id")?.Value,
                Href = item.Attribute("href")?.Value,
                MediaType = item.Attribute("media-type")?.Value
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.Href))
            .ToDictionary(x => x.Id!, x => x, StringComparer.Ordinal);

        var packageDirectory = Path.GetDirectoryName(packagePath)!;
        var ordered = new List<string>();

        foreach (var itemRef in packageDocument.Descendants().Where(x => x.Name.LocalName == "itemref"))
        {
            var idRef = itemRef.Attribute("idref")?.Value;
            if (string.IsNullOrWhiteSpace(idRef) || !manifestById.TryGetValue(idRef, out var manifestItem))
            {
                continue;
            }

            if (!IsHtmlMediaType(manifestItem.MediaType))
            {
                continue;
            }

            ordered.Add(NormalizePackageRelativePath(packageDirectory, manifestItem.Href!));
        }

        foreach (var manifestItem in manifestById.Values.Where(x => IsHtmlMediaType(x.MediaType)))
        {
            var path = NormalizePackageRelativePath(packageDirectory, manifestItem.Href!);
            if (ordered.All(existing => !PathsEqual(existing, path)))
            {
                ordered.Add(path);
            }
        }

        return ordered;
    }

    private static IReadOnlyList<string> ResolveNavigationDocumentPaths(string packagePath, XDocument packageDocument)
    {
        var packageDirectory = Path.GetDirectoryName(packagePath)!;
        return packageDocument
            .Descendants()
            .Where(x => x.Name.LocalName == "item")
            .Where(x => (x.Attribute("properties")?.Value ?? string.Empty).Contains("nav", StringComparison.OrdinalIgnoreCase))
            .Select(x => x.Attribute("href")?.Value)
            .Where(static href => !string.IsNullOrWhiteSpace(href))
            .Select(href => NormalizePackageRelativePath(packageDirectory, href!))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<string> ResolveNcxDocumentPaths(string packagePath, XDocument packageDocument)
    {
        var packageDirectory = Path.GetDirectoryName(packagePath)!;
        return packageDocument
            .Descendants()
            .Where(x => x.Name.LocalName == "item")
            .Where(x => string.Equals(x.Attribute("media-type")?.Value, "application/x-dtbncx+xml", StringComparison.OrdinalIgnoreCase))
            .Select(x => x.Attribute("href")?.Value)
            .Where(static href => !string.IsNullOrWhiteSpace(href))
            .Select(href => NormalizePackageRelativePath(packageDirectory, href!))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsHtmlMediaType(string? mediaType) =>
        string.Equals(mediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(mediaType, "text/html", StringComparison.OrdinalIgnoreCase);

    private static string NormalizePackageRelativePath(string packageDirectory, string relativePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(packageDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(relativePath);

        var decodedRelativePath = Uri.UnescapeDataString(relativePath).Replace('/', Path.DirectorySeparatorChar);
        var resolvedPath = Path.GetFullPath(Path.Combine(packageDirectory, decodedRelativePath));
        var normalizedBaseDirectory = EnsureTrailingDirectorySeparator(Path.GetFullPath(packageDirectory));
        if (!resolvedPath.StartsWith(normalizedBaseDirectory, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"EPUB 资源路径越界，已拒绝访问：{relativePath}");
        }

        return NormalizeDocumentKey(resolvedPath);
    }

    private static string NormalizeAnchorFragment(string fragment)
    {
        if (string.IsNullOrWhiteSpace(fragment))
        {
            return string.Empty;
        }

        return Uri.UnescapeDataString(fragment).Trim().Normalize(NormalizationForm.FormC);
    }

    private static string EnsureTrailingDirectorySeparator(string path) =>
        path.EndsWith(Path.DirectorySeparatorChar) || path.EndsWith(Path.AltDirectorySeparatorChar)
            ? path
            : path + Path.DirectorySeparatorChar;

    private static string NormalizeDocumentKey(string path) =>
        Path.GetFullPath(path).Trim().Normalize(NormalizationForm.FormC);

    private static void PackDirectoryAsEpub(string sourceDirectory, string outputPath)
    {
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Create);

        var mimetypePath = Path.Combine(sourceDirectory, "mimetype");
        if (File.Exists(mimetypePath))
        {
            var mimetypeEntry = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
            using var input = File.OpenRead(mimetypePath);
            using var output = mimetypeEntry.Open();
            input.CopyTo(output);
        }

        foreach (var file in Directory.EnumerateFiles(sourceDirectory, "*", SearchOption.AllDirectories))
        {
            if (PathsEqual(file, mimetypePath))
            {
                continue;
            }

            var relativePath = Path.GetRelativePath(sourceDirectory, file).Replace(Path.DirectorySeparatorChar, '/');
            archive.CreateEntryFromFile(file, relativePath, CompressionLevel.Optimal);
        }
    }

    private static string ResolveTargetExtension(string? value) =>
        string.Equals(value, "DOCX", StringComparison.OrdinalIgnoreCase) ? ".docx" : ".epub";

    private static bool PathsEqual(string left, string right) =>
        string.Equals(NormalizeDocumentKey(left), NormalizeDocumentKey(right), StringComparison.OrdinalIgnoreCase);

    private static string GetWorkingDirectory(string sourcePath)
    {
        var sourceBytes = Encoding.UTF8.GetBytes(Path.GetFullPath(sourcePath));
        var hashBytes = SHA256.HashData(sourceBytes);
        var hash = Convert.ToHexString(hashBytes[..8]).ToLowerInvariant();
        return Path.Combine(
            Path.GetTempPath(),
            "TranslatorApp",
            "ebooks",
            $"{Path.GetFileNameWithoutExtension(sourcePath)}-{hash}");
    }

    private void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (Exception ex)
        {
            Log($"清理电子书临时目录失败：{path}，原因：{ex.Message}");
        }
    }

    private static Task ReportPartialAsync(ITranslationProgressService progressService, string sourcePath, string partial)
    {
        progressService.Publish(Path.GetFileName(sourcePath), partial);
        return Task.CompletedTask;
    }

    private sealed record EpubTranslationFragment(string Original, string ContextHint, string AdditionalRequirements);

    private sealed record EpubTranslationUnit(
        IReadOnlyList<XText> TextNodes,
        string Original,
        string ContextHint,
        string AdditionalRequirements,
        IReadOnlyList<EpubTranslationFragment> Fragments);

    private sealed record EpubContentDocument(
        string Path,
        XDocument Document,
        IReadOnlyList<EpubTranslationUnit> Units,
        string DocumentTitle);

    public sealed record EpubExportDocument(string SourcePath, XDocument Document);

    public sealed record EpubCoverInfo(string? ImagePath, string? DocumentPath);

    public sealed record EpubMetadata(
        string Title,
        IReadOnlyList<string> Creators,
        string Publisher,
        string Language,
        string Date,
        string Identifier,
        string Description);

    private sealed record EpubWorkspace(
        string ExtractDirectory,
        string BookTitle,
        EpubCoverInfo? Cover,
        EpubMetadata Metadata,
        IReadOnlyList<EpubExportDocument> ContentDocuments);

    private sealed record NcxDocument(string Path, int UnitCount);
}
