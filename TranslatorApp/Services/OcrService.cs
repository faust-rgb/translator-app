using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Models;
using Tesseract;
using TranslatorApp.Configuration;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class OcrService(IAppLogService logService) : IOcrService
{
    public Task<IReadOnlyList<OcrTextBlock>> RecognizePdfPageAsync(string pdfPath, int pageIndex, OcrSettings settings, CancellationToken cancellationToken)
    {
        if (!settings.EnableOcrForScannedPdf)
        {
            return Task.FromResult<IReadOnlyList<OcrTextBlock>>(Array.Empty<OcrTextBlock>());
        }

        var tessPath = ResolveTessDataPath(settings.TesseractDataPath);
        if (string.IsNullOrWhiteSpace(tessPath) || !Directory.Exists(tessPath))
        {
            logService.Error("OCR 已启用，但未找到 tessdata 目录。");
            return Task.FromResult<IReadOnlyList<OcrTextBlock>>(Array.Empty<OcrTextBlock>());
        }

        return Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var docLib = DocLib.Instance;
            using var docReader = docLib.GetDocReader(pdfPath, new PageDimensions(settings.RenderScale));
            using var pageReader = docReader.GetPageReader(pageIndex);

            var width = pageReader.GetPageWidth();
            var height = pageReader.GetPageHeight();
            var imageBytes = pageReader.GetImage();
            var pngBytes = ConvertRawBgraToPng(imageBytes, width, height);

            using var engine = new TesseractEngine(tessPath, settings.Language, EngineMode.Default);
            engine.DefaultPageSegMode = PageSegMode.Auto;
            using var pix = Pix.LoadFromMemory(pngBytes);
            using var page = engine.Process(pix);
            using var iterator = page.GetIterator();
            iterator.Begin();

            var words = new List<OcrWord>();
            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                var text = iterator.GetText(PageIteratorLevel.Word);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (iterator.TryGetBoundingBox(PageIteratorLevel.Word, out var rect))
                {
                    var confidence = iterator.GetConfidence(PageIteratorLevel.Word);
                    if (confidence < 35)
                    {
                        continue;
                    }

                    words.Add(new OcrWord(
                        text.Trim(),
                        rect.X1 / settings.RenderScale,
                        rect.Y1 / settings.RenderScale,
                        rect.X2 / settings.RenderScale,
                        rect.Y2 / settings.RenderScale,
                        confidence));
                }
            }
            while (iterator.Next(PageIteratorLevel.Word));

            return BuildBlocks(words, width / settings.RenderScale);
        }, cancellationToken);
    }

    private static IReadOnlyList<OcrTextBlock> BuildBlocks(IReadOnlyList<OcrWord> words, double pageWidth)
    {
        if (words.Count == 0)
        {
            return Array.Empty<OcrTextBlock>();
        }

        var lines = BuildLines(words);
        var columns = BuildColumns(lines, pageWidth);
        var blocks = new List<OcrTextBlock>();

        foreach (var column in columns.OrderBy(x => x.Left))
        {
            var orderedLines = column.Lines.OrderBy(x => x.Top).ToList();
            OcrLine? current = null;

            foreach (var line in orderedLines)
            {
                if (current is null)
                {
                    current = line;
                    continue;
                }

                var verticalGap = line.Top - current.Bottom;
                var indentationGap = Math.Abs(line.Left - current.Left);
                var widthSimilarity = Math.Abs(line.Width - current.Width);
                var shouldMerge =
                    verticalGap < Math.Max(16, current.Height * 0.95) &&
                    indentationGap < Math.Max(18, current.Height * 1.2) &&
                    widthSimilarity < Math.Max(50, current.Width * 0.55);

                if (shouldMerge)
                {
                    current = current.Merge(line);
                }
                else
                {
                    blocks.Add(current.ToBlock());
                    current = line;
                }
            }

            if (current is not null)
            {
                blocks.Add(current.ToBlock());
            }
        }

        return blocks;
    }

    private static List<OcrLine> BuildLines(IReadOnlyList<OcrWord> words)
    {
        var ordered = words.OrderBy(x => x.Top).ThenBy(x => x.Left).ToList();
        var lines = new List<List<OcrWord>>();

        foreach (var word in ordered)
        {
            var target = lines.FirstOrDefault(line =>
            {
                var baseline = line.Average(x => x.Bottom);
                return Math.Abs(baseline - word.Bottom) < Math.Max(6, word.Height * 0.6);
            });

            if (target is null)
            {
                lines.Add([word]);
            }
            else
            {
                target.Add(word);
            }
        }

        return lines
            .Select(line => new OcrLine(line.OrderBy(x => x.Left).ToList()))
            .OrderBy(x => x.Top)
            .ToList();
    }

    private static List<OcrColumn> BuildColumns(IReadOnlyList<OcrLine> lines, double pageWidth)
    {
        var columns = new List<OcrColumn>();

        foreach (var line in lines)
        {
            var column = columns.FirstOrDefault(x =>
                Math.Abs(x.Center - line.Center) < pageWidth * 0.12 ||
                Math.Abs(x.Left - line.Left) < pageWidth * 0.08 ||
                OverlapRatio(x.Left, x.Right, line.Left, line.Right) > 0.35);

            if (column is null)
            {
                columns.Add(new OcrColumn(line));
            }
            else
            {
                column.Add(line);
            }
        }

        return columns;
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

    private static string ResolveTessDataPath(string configuredPath)
    {
        if (!string.IsNullOrWhiteSpace(configuredPath))
        {
            return configuredPath;
        }

        return Path.Combine(AppContext.BaseDirectory, "tessdata");
    }

    private static byte[] ConvertRawBgraToPng(byte[] imageBytes, int width, int height)
    {
        var stride = width * 4;
        var bitmap = BitmapSource.Create(width, height, 96, 96, PixelFormats.Bgra32, null, imageBytes, stride);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));

        using var stream = new MemoryStream();
        encoder.Save(stream);
        return stream.ToArray();
    }

    private sealed record OcrWord(string Text, double Left, double Top, double Right, double Bottom, float Confidence)
    {
        public double Width => Right - Left;
        public double Height => Bottom - Top;
    }

    private sealed class OcrLine
    {
        private readonly IReadOnlyList<OcrWord> _words;

        public OcrLine(IReadOnlyList<OcrWord> words)
        {
            _words = words;
            Text = string.Join(" ", words.Select(x => x.Text));
            Left = words.Min(x => x.Left);
            Right = words.Max(x => x.Right);
            Top = words.Min(x => x.Top);
            Bottom = words.Max(x => x.Bottom);
            Confidence = words.Average(x => x.Confidence);
        }

        public string Text { get; }
        public double Left { get; }
        public double Right { get; }
        public double Top { get; }
        public double Bottom { get; }
        public double Center => (Left + Right) / 2;
        public double Width => Right - Left;
        public double Height => Bottom - Top;
        public double Confidence { get; }

        public OcrLine Merge(OcrLine other) =>
            new(_words.Concat(other._words).OrderBy(x => x.Left).ToList());

        public OcrTextBlock ToBlock() => new(Text, Left, Top, Width, Height);
    }

    private sealed class OcrColumn
    {
        public OcrColumn(OcrLine line)
        {
            Lines = [line];
            Left = line.Left;
            Right = line.Right;
            Center = line.Center;
        }

        public List<OcrLine> Lines { get; }
        public double Left { get; private set; }
        public double Right { get; private set; }
        public double Center { get; private set; }

        public void Add(OcrLine line)
        {
            Lines.Add(line);
            Left = Math.Min(Left, line.Left);
            Right = Math.Max(Right, line.Right);
            Center = (Left + Right) / 2;
        }
    }
}
