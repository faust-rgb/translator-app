using System.IO;
using System.IO.Compression;
using System.Text;
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

        if (context.ResumeUnitIndex > 0)
        {
            Log($"电子书暂不支持从第 {context.ResumeUnitIndex + 1} 个章节继续输出，已自动从头重新生成。");
        }

        var workingDirectory = Path.Combine(
            Path.GetTempPath(),
            "TranslatorApp",
            "ebooks",
            $"{Path.GetFileNameWithoutExtension(context.Item.SourcePath)}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(workingDirectory);

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
        }
        finally
        {
            TryDeleteDirectory(workingDirectory);
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
        ZipFile.ExtractToDirectory(sourceEpubPath, extractDirectory);

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

        var contentDocuments = contentFiles
            .Select((path, index) => new { Path = path, Index = index + 1 })
            .Where(x => IsWithinRequestedRange(x.Index, requestedRange))
            .Select(x => LoadContentDocument(x.Path))
            .Where(static x => x.Units.Count > 0)
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
        var processedUnits = 0;

        foreach (var document in contentDocuments)
        {
            processedUnits = await TranslateContentDocumentAsync(document, context, bilingualSegments, totalUnits, processedUnits);
            document.Document.Save(document.Path, SaveOptions.DisableFormatting);
        }

        var headingMap = BuildHeadingTargetMap(contentDocuments);

        foreach (var navigationDocument in navigationDocuments)
        {
            processedUnits += await TranslateNavigationDocumentAsync(navigationDocument, headingMap, fullContentRangeSelected, context, bilingualSegments);
            var progress = (int)Math.Round(processedUnits * 100d / totalUnits);
            await context.ReportProgressAsync(progress, $"电子书目录 {processedUnits}/{totalUnits}");
            await context.SaveCheckpointAsync(processedUnits, 0, $"电子书目录 {processedUnits}/{totalUnits}");
        }

        foreach (var ncxDocument in ncxDocuments)
        {
            processedUnits += await TranslateNcxDocumentAsync(ncxDocument.Path, headingMap, fullContentRangeSelected, context, bilingualSegments);
            var progress = (int)Math.Round(processedUnits * 100d / totalUnits);
            await context.ReportProgressAsync(progress, $"电子书导航 {processedUnits}/{totalUnits}");
            await context.SaveCheckpointAsync(processedUnits, 0, $"电子书导航 {processedUnits}/{totalUnits}");
        }

        var translatedContentDocuments = contentFiles
            .Select((path, index) => new { Path = path, Index = index + 1 })
            .Where(x => IsWithinRequestedRange(x.Index, requestedRange))
            .Select(x => new EpubExportDocument(x.Path, XDocument.Load(x.Path, LoadOptions.PreserveWhitespace)))
            .ToList();

        return new EpubWorkspace(extractDirectory, bookTitle, cover, metadata, translatedContentDocuments);
    }

    private async Task<int> TranslateContentDocumentAsync(
        EpubContentDocument document,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments,
        int totalUnits,
        int processedUnits)
    {
        var batchSize = GetBlockTranslationConcurrency(context.Settings);
        for (var batchStart = 0; batchStart < document.Units.Count; batchStart += batchSize)
        {
            var batch = document.Units.Skip(batchStart).Take(batchSize).ToList();
            var translatedBatch = await TranslateBatchAsync(
                batch.Select(unit => new TranslationBlock(unit.Original, unit.ContextHint)).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var index = 0; index < batch.Count; index++)
            {
                var unit = batch[index];
                var translated = translatedBatch[index];
                ApplyTranslatedText(unit.TextNodes, translated);
                bilingualSegments.Add(new BilingualSegment(unit.ContextHint, unit.Original, translated));

                processedUnits++;
                var progress = (int)Math.Round(processedUnits * 100d / totalUnits);
                await context.ReportProgressAsync(progress, $"电子书 {Path.GetFileName(document.Path)} {processedUnits}/{totalUnits}");
                await context.SaveCheckpointAsync(processedUnits, 0, $"电子书 {Path.GetFileName(document.Path)} {processedUnits}/{totalUnits}");
            }
        }

        return processedUnits;
    }

    private async Task<int> TranslateNavigationDocumentAsync(
        EpubContentDocument document,
        IReadOnlyDictionary<string, string> headingMap,
        bool allowFallbackTranslation,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments)
    {
        var synchronizedCount = 0;
        var fallbackUnits = new List<(EpubTranslationUnit Unit, string Href)>();
        foreach (var unit in document.Units)
        {
            var href = FindOwningHref(unit.TextNodes);
            var targetKey = string.IsNullOrWhiteSpace(href) ? null : ResolveNavigationTargetKey(document.Path, href!);
            if (targetKey is not null && headingMap.TryGetValue(targetKey, out var synchronizedTitle))
            {
                ApplyTranslatedText(unit.TextNodes, synchronizedTitle);
                bilingualSegments.Add(new BilingualSegment(unit.ContextHint, unit.Original, synchronizedTitle));
                synchronizedCount++;
            }
            else
            {
                fallbackUnits.Add((unit, href ?? string.Empty));
            }
        }

        if (allowFallbackTranslation && fallbackUnits.Count > 0)
        {
            var translatedBatch = await TranslateBatchAsync(
                fallbackUnits.Select(x => new TranslationBlock(x.Unit.Original, x.Unit.ContextHint)).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var index = 0; index < fallbackUnits.Count; index++)
            {
                var unit = fallbackUnits[index].Unit;
                var translated = translatedBatch[index];
                ApplyTranslatedText(unit.TextNodes, translated);
                bilingualSegments.Add(new BilingualSegment(unit.ContextHint, unit.Original, translated));
            }
        }

        document.Document.Save(document.Path, SaveOptions.DisableFormatting);
        return synchronizedCount + (allowFallbackTranslation ? fallbackUnits.Count : 0);
    }

    private async Task<int> TranslateNcxDocumentAsync(
        string ncxPath,
        IReadOnlyDictionary<string, string> headingMap,
        bool allowFallbackTranslation,
        TranslationJobContext context,
        List<BilingualSegment> bilingualSegments)
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

        var fallbackElements = new List<(XElement Element, int Index)>();
        for (var index = 0; index < textElements.Count; index++)
        {
            var element = textElements[index];
            var navPoint = element.Ancestors().FirstOrDefault(x => x.Name.LocalName == "navPoint");
            var src = navPoint?
                .Descendants()
                .FirstOrDefault(x => x.Name.LocalName == "content")
                ?.Attribute("src")
                ?.Value;
            var targetKey = string.IsNullOrWhiteSpace(src) ? null : ResolveNavigationTargetKey(ncxPath, src!);
            if (targetKey is not null && headingMap.TryGetValue(targetKey, out var synchronizedTitle))
            {
                bilingualSegments.Add(new BilingualSegment($"EPUB 目录 {index + 1}", element.Value, synchronizedTitle));
                element.Value = synchronizedTitle;
            }
            else
            {
                fallbackElements.Add((element, index));
            }
        }

        if (allowFallbackTranslation && fallbackElements.Count > 0)
        {
            var translatedBatch = await TranslateBatchAsync(
                fallbackElements.Select(x => new TranslationBlock(x.Element.Value, $"EPUB 目录 {x.Index + 1}")).ToList(),
                context.Settings,
                context.PauseController,
                partial => ReportPartialAsync(progressService, context.Item.SourcePath, partial),
                context.CancellationToken);

            for (var index = 0; index < fallbackElements.Count; index++)
            {
                var element = fallbackElements[index].Element;
                bilingualSegments.Add(new BilingualSegment($"EPUB 目录 {fallbackElements[index].Index + 1}", element.Value, translatedBatch[index]));
                element.Value = translatedBatch[index];
            }
        }

        document.Save(ncxPath, SaveOptions.DisableFormatting);
        return textElements.Count;
    }

    private static Dictionary<string, string> BuildHeadingTargetMap(IReadOnlyList<EpubContentDocument> contentDocuments)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var document in contentDocuments)
        {
            var headings = document.Document
                .Descendants()
                .Where(IsHeadingElement)
                .ToList();
            var firstHeading = headings
                .Select(GetElementText)
                .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));

            if (!string.IsNullOrWhiteSpace(firstHeading))
            {
                map[Path.GetFullPath(document.Path)] = firstHeading!;
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
                    map[$"{Path.GetFullPath(document.Path)}#{id}"] = text;
                }
            }
        }

        return map;
    }

    private static bool IsHeadingElement(XElement element) =>
        element.Name.LocalName is "h1" or "h2" or "h3" or "h4" or "h5" or "h6";

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

        var hashIndex = href.IndexOf('#');
        var filePart = hashIndex >= 0 ? href[..hashIndex] : href;
        var fragment = hashIndex >= 0 ? href[(hashIndex + 1)..] : string.Empty;
        var baseDirectory = Path.GetDirectoryName(navigationDocumentPath)!;
        var targetPath = string.IsNullOrWhiteSpace(filePart)
            ? Path.GetFullPath(navigationDocumentPath)
            : NormalizePackageRelativePath(baseDirectory, filePart);
        var normalizedFragment = NormalizeAnchorFragment(fragment);

        return string.IsNullOrWhiteSpace(normalizedFragment)
            ? targetPath
            : $"{targetPath}#{normalizedFragment}";
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
        var units = BuildTranslationUnits(document, Path.GetFileName(path));
        return new EpubContentDocument(path, document, units);
    }

    private static int GetNcxUnitCount(string path)
    {
        var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        return document
            .Descendants()
            .Count(x => x.Name.LocalName == "text" && !string.IsNullOrWhiteSpace(x.Value));
    }

    private static List<EpubTranslationUnit> BuildTranslationUnits(XDocument document, string fileName)
    {
        var units = new List<EpubTranslationUnit>();
        var candidates = document
            .Descendants()
            .Where(IsTranslatableLeafElement)
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

            units.Add(new EpubTranslationUnit(textNodes, original, $"EPUB {fileName} <{element.Name.LocalName}>"));
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

        return resolvedPath;
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
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);

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

    private sealed record EpubTranslationUnit(IReadOnlyList<XText> TextNodes, string Original, string ContextHint);

    private sealed record EpubContentDocument(string Path, XDocument Document, IReadOnlyList<EpubTranslationUnit> Units);

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
