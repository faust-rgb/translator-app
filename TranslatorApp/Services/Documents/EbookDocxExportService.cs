using System.Globalization;
using System.IO;
using System.Xml.Linq;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace TranslatorApp.Services.Documents;

public sealed class EbookDocxExportService : IEbookDocxExportService
{
    private const long EmusPerPixel = 9525;
    private const long MaxImageWidthEmus = 6_000_000;
    private const uint DefaultImageWidthPixels = 480;
    private const uint DefaultImageHeightPixels = 320;

    public Task ExportAsync(
        string outputPath,
        string title,
        EbookDocumentTranslator.EpubCoverInfo? cover,
        EbookDocumentTranslator.EpubMetadata metadata,
        IReadOnlyList<EbookDocumentTranslator.EpubExportDocument> contentDocuments,
        CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        using var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;
        var context = new ExportContext(mainPart);
        EnsureUpdateFieldsOnOpen(document);
        var coverDocument = ResolveCoverDocument(cover, contentDocuments);

        var hasCover = AppendCoverPage(body, title, cover, coverDocument, metadata, context);
        if (hasCover)
        {
            body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        AppendMetadataPage(body, metadata);
        AppendDocxToc(body);
        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

        foreach (var contentDocument in contentDocuments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (coverDocument is not null &&
                string.Equals(Path.GetFullPath(contentDocument.SourcePath), Path.GetFullPath(coverDocument.SourcePath), StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            AppendContentDocument(body, contentDocument, context, cancellationToken);
        }

        mainPart.Document.Save();
        return Task.CompletedTask;
    }

    private static bool AppendCoverPage(
        Body body,
        string title,
        EbookDocumentTranslator.EpubCoverInfo? cover,
        EbookDocumentTranslator.EpubExportDocument? coverDocument,
        EbookDocumentTranslator.EpubMetadata metadata,
        ExportContext context)
    {
        var hasContent = false;

        if (cover?.ImagePath is not null && File.Exists(cover.ImagePath))
        {
            var imageElement = new XElement("img", new XAttribute("src", cover.ImagePath));
            var drawing = TryCreateImageDrawing(imageElement, context, cover.ImagePath, resourceAlreadyResolved: true);
            if (drawing is not null)
            {
                body.Append(new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                    new Run(drawing)));
                hasContent = true;
            }
        }

        var coverTexts = ExtractCoverTexts(coverDocument?.Document);
        if (coverTexts.Count == 0 && !string.IsNullOrWhiteSpace(title))
        {
            coverTexts.Add(title);
        }

        for (var index = 0; index < coverTexts.Count; index++)
        {
            var styleId = index == 0 ? "Title" : "Subtitle";
            body.Append(new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = styleId },
                    new Justification { Val = JustificationValues.Center },
                    new SpacingBetweenLines
                    {
                        Before = index == 0 ? "240" : "60",
                        After = index == 0 ? "180" : "60"
                    }),
                new Run(new Text(coverTexts[index]) { Space = SpaceProcessingModeValues.Preserve })));
            hasContent = true;
        }

        foreach (var line in BuildCoverMetadataLines(metadata))
        {
            body.Append(new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center },
                    new SpacingBetweenLines { Before = "30", After = "30" }),
                new Run(
                    new RunProperties(new Color { Val = "666666" }),
                    new Text(line) { Space = SpaceProcessingModeValues.Preserve })));
            hasContent = true;
        }

        return hasContent;
    }

    private static void AppendDocxToc(Body body)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("目录"))));

        body.Append(new Paragraph(
            new SimpleField
            {
                Instruction = @"TOC \o ""1-6"" \h \z \u"
            }));
    }

    private static void AppendMetadataPage(Body body, EbookDocumentTranslator.EpubMetadata metadata)
    {
        var lines = new List<string>();
        if (!string.IsNullOrWhiteSpace(metadata.Description))
        {
            lines.Add("内容简介");
            lines.Add(metadata.Description.Trim());
        }

        if (!string.IsNullOrWhiteSpace(metadata.Identifier))
        {
            lines.Add($"书籍标识：{metadata.Identifier}");
        }

        if (lines.Count == 0)
        {
            return;
        }

        foreach (var line in lines)
        {
            var isHeading = string.Equals(line, "内容简介", StringComparison.Ordinal);
            body.Append(new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = isHeading ? "Heading1" : "Normal" },
                    new SpacingBetweenLines
                    {
                        Before = isHeading ? "120" : "30",
                        After = isHeading ? "90" : "30"
                    }),
                new Run(new Text(line) { Space = SpaceProcessingModeValues.Preserve })));
        }

        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
    }

    private static void EnsureUpdateFieldsOnOpen(WordprocessingDocument document)
    {
        var settingsPart = document.MainDocumentPart!.DocumentSettingsPart ?? document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        var updateFields = settingsPart.Settings.Elements<UpdateFieldsOnOpen>().FirstOrDefault();
        if (updateFields is null)
        {
            settingsPart.Settings.Append(new UpdateFieldsOnOpen { Val = true });
        }
        else
        {
            updateFields.Val = true;
        }

        settingsPart.Settings.Save();
    }

    private static EbookDocumentTranslator.EpubExportDocument? ResolveCoverDocument(
        EbookDocumentTranslator.EpubCoverInfo? cover,
        IReadOnlyList<EbookDocumentTranslator.EpubExportDocument> contentDocuments)
    {
        if (cover?.DocumentPath is null)
        {
            return null;
        }

        return contentDocuments.FirstOrDefault(document =>
            string.Equals(Path.GetFullPath(document.SourcePath), Path.GetFullPath(cover.DocumentPath), StringComparison.OrdinalIgnoreCase));
    }

    private static List<string> ExtractCoverTexts(XDocument? document)
    {
        if (document?.Root is null)
        {
            return [];
        }

        var bodyElement = document.Root.Descendants().FirstOrDefault(x => x.Name.LocalName == "body") ?? document.Root;
        var texts = new List<string>();
        foreach (var element in bodyElement.Descendants().Where(IsCoverTextCandidate))
        {
            var text = string.Concat(element.DescendantNodes().OfType<XText>().Select(x => x.Value)).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (texts.Any(existing => string.Equals(existing, text, StringComparison.Ordinal)))
            {
                continue;
            }

            texts.Add(text);
            if (texts.Count >= 4)
            {
                break;
            }
        }

        return texts;
    }

    private static bool IsCoverTextCandidate(XElement element) =>
        element.Name.LocalName is "h1" or "h2" or "h3" or "h4" or "p" or "div";

    private static List<string> BuildCoverMetadataLines(EbookDocumentTranslator.EpubMetadata metadata)
    {
        var lines = new List<string>();

        if (metadata.Creators.Count > 0)
        {
            lines.Add($"作者：{string.Join(" / ", metadata.Creators)}");
        }

        if (!string.IsNullOrWhiteSpace(metadata.Publisher))
        {
            lines.Add($"出版社：{metadata.Publisher}");
        }

        var facts = new List<string>();
        if (!string.IsNullOrWhiteSpace(metadata.Language))
        {
            facts.Add($"语言：{metadata.Language}");
        }

        if (!string.IsNullOrWhiteSpace(metadata.Date))
        {
            facts.Add($"日期：{metadata.Date}");
        }

        if (!string.IsNullOrWhiteSpace(metadata.Identifier))
        {
            facts.Add($"标识：{metadata.Identifier}");
        }

        if (facts.Count > 0)
        {
            lines.Add(string.Join("    ", facts));
        }

        return lines;
    }

    private static void AppendContentDocument(
        Body body,
        EbookDocumentTranslator.EpubExportDocument contentDocument,
        ExportContext context,
        CancellationToken cancellationToken)
    {
        var root = contentDocument.Document.Root;
        if (root is null)
        {
            return;
        }

        var bodyElement = root.Descendants().FirstOrDefault(x => x.Name.LocalName == "body") ?? root;
        var listState = new Stack<ListContext>();
        foreach (var child in bodyElement.Nodes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            AppendBlockNode(body, child, listState, context, contentDocument.SourcePath, cancellationToken);
        }
    }

    private static void AppendBlockNode(
        Body body,
        XNode node,
        Stack<ListContext> listState,
        ExportContext context,
        string sourcePath,
        CancellationToken cancellationToken)
    {
        if (node is XElement element)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var name = element.Name.LocalName.ToLowerInvariant();
            switch (name)
            {
                case "section":
                case "article":
                case "main":
                case "body":
                case "div":
                    AppendContainer(body, element, listState, context, sourcePath, cancellationToken);
                    break;
                case "figure":
                    AppendFigure(body, element, context, sourcePath);
                    break;
                case "p":
                case "blockquote":
                case "figcaption":
                case "caption":
                case "dt":
                case "dd":
                    AppendParagraph(body, element, CreateParagraphProperties(element, listState, false), listState, context, sourcePath);
                    break;
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    AppendParagraph(body, element, CreateHeadingProperties(name, element), listState, context, sourcePath);
                    break;
                case "ul":
                    listState.Push(new ListContext(false));
                    AppendContainer(body, element, listState, context, sourcePath, cancellationToken);
                    listState.Pop();
                    break;
                case "ol":
                    listState.Push(new ListContext(true));
                    AppendContainer(body, element, listState, context, sourcePath, cancellationToken);
                    listState.Pop();
                    break;
                case "li":
                    AppendParagraph(body, element, CreateParagraphProperties(element, listState, true), listState, context, sourcePath);
                    break;
                case "table":
                    AppendTable(body, element, context, sourcePath);
                    break;
                case "img":
                    AppendStandaloneImage(body, element, context, sourcePath);
                    break;
                default:
                    AppendContainer(body, element, listState, context, sourcePath, cancellationToken);
                    break;
            }
        }
        else if (node is XText text && !string.IsNullOrWhiteSpace(text.Value))
        {
            body.Append(new Paragraph(new Run(new Text(text.Value) { Space = SpaceProcessingModeValues.Preserve })));
        }
    }

    private static void AppendContainer(
        Body body,
        XElement element,
        Stack<ListContext> listState,
        ExportContext context,
        string sourcePath,
        CancellationToken cancellationToken)
    {
        foreach (var child in element.Nodes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            AppendBlockNode(body, child, listState, context, sourcePath, cancellationToken);
        }
    }

    private static void AppendParagraph(
        Body body,
        XElement element,
        ParagraphProperties properties,
        Stack<ListContext> listState,
        ExportContext context,
        string sourcePath)
    {
        var paragraph = CreateParagraphFromElement(element, properties, context, sourcePath);

        if (element.Name.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase) && listState.Count > 0)
        {
            paragraph.InsertAt(new Run(new Text(listState.Peek().NextMarker()) { Space = SpaceProcessingModeValues.Preserve }), 1);
        }

        if (!paragraph.Elements<Run>().Any())
        {
            return;
        }

        body.Append(paragraph);
    }

    private static Paragraph CreateParagraphFromElement(
        XElement element,
        ParagraphProperties properties,
        ExportContext context,
        string sourcePath)
    {
        var paragraph = new Paragraph();
        paragraph.Append(properties.CloneNode(true));

        foreach (var run in BuildRuns(element.Nodes(), new InlineStyle(), context, sourcePath))
        {
            paragraph.Append(run);
        }

        return paragraph;
    }

    private static ParagraphProperties CreateHeadingProperties(string headingName, XElement element)
    {
        var level = headingName switch
        {
            "h1" => 1,
            "h2" => 2,
            "h3" => 3,
            "h4" => 4,
            "h5" => 5,
            _ => 6
        };

        var properties = new ParagraphProperties(
            new ParagraphStyleId { Val = $"Heading{level}" },
            new SpacingBetweenLines { Before = "240", After = "120" });

        ApplyCommonParagraphStyle(properties, element, new Stack<ListContext>(), isListItem: false);
        return properties;
    }

    private static ParagraphProperties CreateParagraphProperties(XElement element, Stack<ListContext> listState, bool isListItem)
    {
        var properties = new ParagraphProperties(
            new SpacingBetweenLines { After = "90", Line = "300", LineRule = LineSpacingRuleValues.Auto });

        if (element.Name.LocalName.Equals("blockquote", StringComparison.OrdinalIgnoreCase))
        {
            properties.Append(new Indentation { Left = "720", Right = "240" });
        }

        ApplyCommonParagraphStyle(properties, element, listState, isListItem);
        return properties;
    }

    private static void ApplyCommonParagraphStyle(ParagraphProperties properties, XElement element, Stack<ListContext> listState, bool isListItem)
    {
        var style = ParseCssStyle(element.Attribute("style")?.Value);
        var className = element.Attribute("class")?.Value ?? string.Empty;

        var alignment = ResolveJustification(style, className);
        if (alignment is not null)
        {
            properties.Append(new Justification { Val = alignment.Value });
        }

        var indentation = new Indentation();
        var hasIndentation = false;

        if (isListItem && listState.Count > 0)
        {
            var indent = 360 * listState.Count;
            indentation.Left = indent.ToString(CultureInfo.InvariantCulture);
            indentation.Hanging = "240";
            hasIndentation = true;
        }

        if (TryParseCssLength(style, "margin-left", out var marginLeftTwips))
        {
            indentation.Left = marginLeftTwips.ToString(CultureInfo.InvariantCulture);
            hasIndentation = true;
        }

        if (TryParseCssLength(style, "margin-right", out var marginRightTwips))
        {
            indentation.Right = marginRightTwips.ToString(CultureInfo.InvariantCulture);
            hasIndentation = true;
        }

        if (TryParseCssLength(style, "text-indent", out var textIndentTwips))
        {
            indentation.FirstLine = textIndentTwips.ToString(CultureInfo.InvariantCulture);
            hasIndentation = true;
        }

        if (hasIndentation)
        {
            properties.Append(indentation);
        }

        var hasTopMargin = TryParseCssLength(style, "margin-top", out var topTwips);
        var hasBottomMargin = TryParseCssLength(style, "margin-bottom", out var bottomTwips);
        if (hasTopMargin || hasBottomMargin)
        {
            properties.Append(new SpacingBetweenLines
            {
                Before = hasTopMargin && topTwips > 0 ? topTwips.ToString(CultureInfo.InvariantCulture) : null,
                After = hasBottomMargin && bottomTwips > 0 ? bottomTwips.ToString(CultureInfo.InvariantCulture) : null,
                Line = "300",
                LineRule = LineSpacingRuleValues.Auto
            });
        }
    }

    private static void AppendTable(Body body, XElement tableElement, ExportContext context, string sourcePath)
    {
        var rows = tableElement
            .Descendants()
            .Where(x => x.Name.LocalName.Equals("tr", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (rows.Count == 0)
        {
            return;
        }

        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        foreach (var rowElement in rows)
        {
            var row = new TableRow();
            var cells = rowElement.Elements().Where(x =>
                x.Name.LocalName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
                x.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase));

            foreach (var cellElement in cells)
            {
                var cellParagraph = new Paragraph();
                foreach (var run in BuildRuns(cellElement.Nodes(), new InlineStyle(), context, sourcePath))
                {
                    cellParagraph.Append(run);
                }

                if (!cellParagraph.Elements<Run>().Any())
                {
                    cellParagraph.Append(new Run(new Text(cellElement.Value ?? string.Empty)));
                }

                var cell = new TableCell(cellParagraph);
                row.Append(cell);
            }

            if (row.Elements<TableCell>().Any())
            {
                table.Append(row);
            }
        }

        if (table.Elements<TableRow>().Any())
        {
            body.Append(table);
        }
    }

    private static void AppendStandaloneImage(Body body, XElement imageElement, ExportContext context, string sourcePath)
    {
        var paragraph = new Paragraph(CreateParagraphProperties(imageElement, new Stack<ListContext>(), false));
        var drawing = TryCreateImageDrawing(imageElement, context, sourcePath);
        if (drawing is not null)
        {
            paragraph.Append(new Run(drawing));
        }
        else
        {
            paragraph.Append(CreateImageFallbackRun(imageElement));
        }

        body.Append(paragraph);
    }

    private static void AppendFigure(Body body, XElement figureElement, ExportContext context, string sourcePath)
    {
        var imageElement = figureElement.Descendants().FirstOrDefault(x => x.Name.LocalName == "img");
        var captionElement = figureElement.Descendants().FirstOrDefault(x => x.Name.LocalName == "figcaption");

        if (imageElement is not null)
        {
            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center },
                    new SpacingBetweenLines { Before = "60", After = "60" }));
            var drawing = TryCreateImageDrawing(imageElement, context, sourcePath);
            paragraph.Append(drawing is not null ? new Run(drawing) : CreateImageFallbackRun(imageElement));
            body.Append(paragraph);
        }

        if (captionElement is not null)
        {
            var properties = CreateParagraphProperties(captionElement, new Stack<ListContext>(), false);
            properties.Append(new Justification { Val = JustificationValues.Center });
            var paragraph = CreateParagraphFromElement(captionElement, properties, context, sourcePath);
            if (paragraph.Elements<Run>().Any())
            {
                body.Append(paragraph);
            }
        }
    }

    private static IEnumerable<Run> BuildRuns(
        IEnumerable<XNode> nodes,
        InlineStyle inheritedStyle,
        ExportContext context,
        string sourcePath)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case XText textNode:
                    if (!string.IsNullOrEmpty(textNode.Value))
                    {
                        yield return CreateRun(textNode.Value, inheritedStyle);
                    }
                    break;
                case XElement element:
                    var name = element.Name.LocalName.ToLowerInvariant();
                    if (name == "br")
                    {
                        yield return new Run(new Break());
                        continue;
                    }

                    if (name == "img")
                    {
                        var drawing = TryCreateImageDrawing(element, context, sourcePath);
                        yield return drawing is not null
                            ? new Run(drawing)
                            : CreateImageFallbackRun(element);
                        continue;
                    }

                    if (IsBlockElement(name))
                    {
                        continue;
                    }

                    var style = MergeInlineStyle(inheritedStyle, element);
                    foreach (var run in BuildRuns(element.Nodes(), style, context, sourcePath))
                    {
                        yield return run;
                    }
                    break;
            }
        }
    }

    private static Run CreateRun(string text, InlineStyle style)
    {
        var properties = new RunProperties();
        if (style.Bold)
        {
            properties.Append(new Bold());
        }

        if (style.Italic)
        {
            properties.Append(new Italic());
        }

        if (style.Underline)
        {
            properties.Append(new Underline { Val = UnderlineValues.Single });
        }

        if (style.Subscript)
        {
            properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        }

        if (style.Superscript)
        {
            properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        }

        if (style.Code)
        {
            properties.Append(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" });
        }

        if (style.FontSizeHalfPoints is not null)
        {
            properties.Append(new FontSize { Val = style.FontSizeHalfPoints.Value.ToString(CultureInfo.InvariantCulture) });
        }

        if (!string.IsNullOrWhiteSpace(style.ColorHex))
        {
            properties.Append(new Color { Val = style.ColorHex });
        }

        var textNode = new Text(text) { Space = SpaceProcessingModeValues.Preserve };
        return properties.ChildElements.Count > 0
            ? new Run(properties, textNode)
            : new Run(textNode);
    }

    private static InlineStyle MergeInlineStyle(InlineStyle inheritedStyle, XElement element)
    {
        var style = inheritedStyle;
        var name = element.Name.LocalName.ToLowerInvariant();

        if (name is "strong" or "b")
        {
            style = style with { Bold = true };
        }

        if (name is "em" or "i" or "cite")
        {
            style = style with { Italic = true };
        }

        if (name == "u")
        {
            style = style with { Underline = true };
        }

        if (name == "sub")
        {
            style = style with { Subscript = true, Superscript = false };
        }

        if (name == "sup")
        {
            style = style with { Superscript = true, Subscript = false };
        }

        if (name is "code" or "kbd" or "samp")
        {
            style = style with { Code = true };
        }

        if (name == "a")
        {
            style = style with { Underline = true };
        }

        var className = element.Attribute("class")?.Value ?? string.Empty;
        if (className.Contains("bold", StringComparison.OrdinalIgnoreCase))
        {
            style = style with { Bold = true };
        }

        if (className.Contains("italic", StringComparison.OrdinalIgnoreCase))
        {
            style = style with { Italic = true };
        }

        var css = ParseCssStyle(element.Attribute("style")?.Value);
        if (css.TryGetValue("font-weight", out var fontWeight) &&
            (fontWeight.Contains("bold", StringComparison.OrdinalIgnoreCase) || fontWeight == "700"))
        {
            style = style with { Bold = true };
        }

        if (css.TryGetValue("font-style", out var fontStyle) &&
            fontStyle.Contains("italic", StringComparison.OrdinalIgnoreCase))
        {
            style = style with { Italic = true };
        }

        if (css.TryGetValue("text-decoration", out var textDecoration) &&
            textDecoration.Contains("underline", StringComparison.OrdinalIgnoreCase))
        {
            style = style with { Underline = true };
        }

        if (TryParseCssLength(css, "font-size", out var fontSizeTwips))
        {
            var halfPoints = Math.Max(16, (int)Math.Round(fontSizeTwips / 10d));
            style = style with { FontSizeHalfPoints = halfPoints };
        }

        if (css.TryGetValue("color", out var colorHex) && TryNormalizeColor(colorHex, out var normalizedColor))
        {
            style = style with { ColorHex = normalizedColor };
        }

        return style;
    }

    private static Drawing? TryCreateImageDrawing(XElement imageElement, ExportContext context, string sourcePath, bool resourceAlreadyResolved = false)
    {
        var src = imageElement.Attribute("src")?.Value;
        if (string.IsNullOrWhiteSpace(src) || src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var imagePath = resourceAlreadyResolved ? src : ResolveResourcePath(sourcePath, src);
        if (!File.Exists(imagePath))
        {
            return null;
        }

        var imagePartType = TryResolveImagePartType(Path.GetExtension(imagePath));
        if (imagePartType is null)
        {
            return null;
        }

        if (!context.ImageRelationships.TryGetValue(imagePath, out var relationshipId))
        {
            var imagePart = context.MainPart.AddImagePart(imagePartType.Value);
            using var stream = File.OpenRead(imagePath);
            imagePart.FeedData(stream);
            relationshipId = context.MainPart.GetIdOfPart(imagePart);
            context.ImageRelationships[imagePath] = relationshipId;
        }

        var (widthEmus, heightEmus) = ResolveImageSize(imageElement);
        var name = Path.GetFileName(imagePath);
        var docPrId = context.NextDrawingId++;

        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmus, Cy = heightEmus },
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties { Id = docPrId, Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = docPrId, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmus, Cy = heightEmus }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });
    }

    private static Run CreateImageFallbackRun(XElement imageElement)
    {
        var alt = imageElement.Attribute("alt")?.Value;
        if (string.IsNullOrWhiteSpace(alt))
        {
            alt = imageElement.Attribute(XName.Get("alt", "http://www.w3.org/XML/1998/namespace"))?.Value;
        }

        var text = string.IsNullOrWhiteSpace(alt) ? "[图片]" : $"[图片] {alt}";
        return new Run(
            new RunProperties(new Italic()),
            new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static string ResolveResourcePath(string sourcePath, string relativeResourcePath)
    {
        var baseDirectory = Path.GetDirectoryName(sourcePath)!;
        var normalizedPath = relativeResourcePath.Replace('/', Path.DirectorySeparatorChar);
        var resolvedPath = Path.GetFullPath(Path.Combine(baseDirectory, normalizedPath));
        var normalizedBaseDirectory = EnsureTrailingDirectorySeparator(Path.GetFullPath(baseDirectory));
        if (!resolvedPath.StartsWith(normalizedBaseDirectory, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"资源路径越界，已拒绝访问：{relativeResourcePath}");
        }

        return resolvedPath;
    }

    private static string EnsureTrailingDirectorySeparator(string path) =>
        path.EndsWith(Path.DirectorySeparatorChar) || path.EndsWith(Path.AltDirectorySeparatorChar)
            ? path
            : path + Path.DirectorySeparatorChar;

    private static PartTypeInfo? TryResolveImagePartType(string extension) =>
        extension.ToLowerInvariant() switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            ".ico" => ImagePartType.Icon,
            ".svg" => ImagePartType.Svg,
            _ => null
        };

    private static (long WidthEmus, long HeightEmus) ResolveImageSize(XElement imageElement)
    {
        var style = ParseCssStyle(imageElement.Attribute("style")?.Value);
        long widthEmus = 0;
        long heightEmus = 0;
        (uint Width, uint Height)? intrinsicSize = null;
        var src = imageElement.Attribute("src")?.Value;
        if (!string.IsNullOrWhiteSpace(src) && File.Exists(src))
        {
            intrinsicSize = TryReadImagePixelSize(src);
        }

        if (TryParseHtmlPixels(imageElement.Attribute("width")?.Value, out var htmlWidth))
        {
            widthEmus = htmlWidth * EmusPerPixel;
        }

        if (TryParseHtmlPixels(imageElement.Attribute("height")?.Value, out var htmlHeight))
        {
            heightEmus = htmlHeight * EmusPerPixel;
        }

        if (TryParseCssPixels(style, "width", out var cssWidth))
        {
            widthEmus = cssWidth * EmusPerPixel;
        }

        if (TryParseCssPixels(style, "height", out var cssHeight))
        {
            heightEmus = cssHeight * EmusPerPixel;
        }

        if (widthEmus == 0 && heightEmus == 0 && intrinsicSize is not null)
        {
            widthEmus = intrinsicSize.Value.Width * EmusPerPixel;
            heightEmus = intrinsicSize.Value.Height * EmusPerPixel;
        }
        else if (widthEmus > 0 && heightEmus == 0 && intrinsicSize is not null && intrinsicSize.Value.Width > 0)
        {
            heightEmus = (long)Math.Round(widthEmus * (intrinsicSize.Value.Height / (double)intrinsicSize.Value.Width));
        }
        else if (heightEmus > 0 && widthEmus == 0 && intrinsicSize is not null && intrinsicSize.Value.Height > 0)
        {
            widthEmus = (long)Math.Round(heightEmus * (intrinsicSize.Value.Width / (double)intrinsicSize.Value.Height));
        }

        if (widthEmus == 0)
        {
            widthEmus = DefaultImageWidthPixels * EmusPerPixel;
        }

        if (heightEmus == 0)
        {
            heightEmus = DefaultImageHeightPixels * EmusPerPixel;
        }

        if (widthEmus > MaxImageWidthEmus && widthEmus > 0)
        {
            var scale = MaxImageWidthEmus / (double)widthEmus;
            widthEmus = MaxImageWidthEmus;
            heightEmus = Math.Max(EmusPerPixel * 32, (long)Math.Round(heightEmus * scale));
        }

        return (widthEmus, heightEmus);
    }

    private static (uint Width, uint Height)? TryReadImagePixelSize(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
            var frame = decoder.Frames.FirstOrDefault();
            if (frame is null || frame.PixelWidth <= 0 || frame.PixelHeight <= 0)
            {
                return null;
            }

            return ((uint)frame.PixelWidth, (uint)frame.PixelHeight);
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<string, string> ParseCssStyle(string? styleValue)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(styleValue))
        {
            return result;
        }

        foreach (var item in styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var separatorIndex = item.IndexOf(':');
            if (separatorIndex <= 0)
            {
                continue;
            }

            var key = item[..separatorIndex].Trim();
            var value = item[(separatorIndex + 1)..].Trim();
            if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
            {
                result[key] = value;
            }
        }

        return result;
    }

    private static bool TryParseCssLength(IReadOnlyDictionary<string, string> css, string name, out int twips)
    {
        twips = 0;
        if (!css.TryGetValue(name, out var value))
        {
            return false;
        }

        value = value.Trim().ToLowerInvariant();
        if (value.EndsWith("px", StringComparison.Ordinal))
        {
            if (double.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var pixels))
            {
                twips = Math.Max(0, (int)Math.Round(pixels * 15));
                return true;
            }
        }

        if (value.EndsWith("pt", StringComparison.Ordinal))
        {
            if (double.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var points))
            {
                twips = Math.Max(0, (int)Math.Round(points * 20));
                return true;
            }
        }

        if (value.EndsWith("em", StringComparison.Ordinal))
        {
            if (double.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var em))
            {
                twips = Math.Max(0, (int)Math.Round(em * 240));
                return true;
            }
        }

        return false;
    }

    private static bool TryParseCssPixels(IReadOnlyDictionary<string, string> css, string name, out long pixels)
    {
        pixels = 0;
        if (!css.TryGetValue(name, out var value))
        {
            return false;
        }

        return TryParseHtmlPixels(value, out pixels);
    }

    private static bool TryParseHtmlPixels(string? value, out long pixels)
    {
        pixels = 0;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        value = value.Trim().ToLowerInvariant();
        if (value.EndsWith("px", StringComparison.Ordinal))
        {
            value = value[..^2];
        }

        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
        {
            return false;
        }

        pixels = Math.Max(1, (long)Math.Round(parsed));
        return true;
    }

    private static JustificationValues? ResolveJustification(IReadOnlyDictionary<string, string> css, string className)
    {
        if (css.TryGetValue("text-align", out var textAlign))
        {
            return textAlign.Trim().ToLowerInvariant() switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                "justify" => JustificationValues.Both,
                _ => JustificationValues.Left
            };
        }

        if (className.Contains("center", StringComparison.OrdinalIgnoreCase))
        {
            return JustificationValues.Center;
        }

        if (className.Contains("right", StringComparison.OrdinalIgnoreCase))
        {
            return JustificationValues.Right;
        }

        return null;
    }

    private static bool TryNormalizeColor(string value, out string normalizedColor)
    {
        normalizedColor = string.Empty;
        value = value.Trim();

        if (value.StartsWith('#'))
        {
            var hex = value[1..];
            if (hex.Length == 3)
            {
                normalizedColor = string.Concat(hex.Select(ch => $"{ch}{ch}")).ToUpperInvariant();
                return true;
            }

            if (hex.Length == 6)
            {
                normalizedColor = hex.ToUpperInvariant();
                return true;
            }
        }

        return value.ToLowerInvariant() switch
        {
            "black" => SetColor("000000", out normalizedColor),
            "white" => SetColor("FFFFFF", out normalizedColor),
            "red" => SetColor("FF0000", out normalizedColor),
            "blue" => SetColor("0000FF", out normalizedColor),
            "green" => SetColor("008000", out normalizedColor),
            "gray" or "grey" => SetColor("808080", out normalizedColor),
            _ => false
        };
    }

    private static bool SetColor(string value, out string normalizedColor)
    {
        normalizedColor = value;
        return true;
    }

    private static bool IsBlockElement(string name) =>
        name is "p" or "div" or "section" or "article" or "main" or "ul" or "ol" or "li" or "table" or "tr" or "td" or "th" or
            "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "blockquote" or "figure" or "figcaption";

    private sealed record ListContext(bool Ordered)
    {
        public int Counter { get; private set; }

        public string NextMarker()
        {
            if (!Ordered)
            {
                return "\u2022 ";
            }

            Counter++;
            return $"{Counter}. ";
        }
    }

    private sealed record ExportContext(MainDocumentPart MainPart)
    {
        public Dictionary<string, string> ImageRelationships { get; } = new(StringComparer.OrdinalIgnoreCase);

        public uint NextDrawingId { get; set; } = 1;
    }

    private readonly record struct InlineStyle(
        bool Bold = false,
        bool Italic = false,
        bool Underline = false,
        bool Superscript = false,
        bool Subscript = false,
        bool Code = false,
        int? FontSizeHalfPoints = null,
        string? ColorHex = null);
}
