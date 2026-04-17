using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var arguments = CommandLineArguments.Parse(args);
var jsonOptions = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};
if (arguments.ShowHelp)
{
    CommandLineArguments.PrintUsage();
    return 0;
}

if (arguments.ErrorMessage is not null)
{
    Console.Error.WriteLine(arguments.ErrorMessage);
    CommandLineArguments.PrintUsage();
    return 1;
}

if (!File.Exists(arguments.PdfPath))
{
    Console.Error.WriteLine($"找不到 PDF 文件：{arguments.PdfPath}");
    return 1;
}

if (!File.Exists(arguments.BilingualPath))
{
    Console.Error.WriteLine($"找不到双语 DOCX 文件：{arguments.BilingualPath}");
    return 1;
}

var report = Inspector.Run(arguments.PdfPath, arguments.BilingualPath);
Directory.CreateDirectory(Path.GetDirectoryName(arguments.ReportTextPath)!);
Directory.CreateDirectory(Path.GetDirectoryName(arguments.ReportJsonPath)!);

await File.WriteAllTextAsync(arguments.ReportTextPath, ReportWriter.WriteText(report), Encoding.UTF8);
await File.WriteAllTextAsync(
    arguments.ReportJsonPath,
    JsonSerializer.Serialize(report, jsonOptions),
    Encoding.UTF8);

Console.WriteLine($"PDF: {report.PdfPath}");
Console.WriteLine($"双语 DOCX: {report.BilingualPath}");
Console.WriteLine($"总片段数: {report.TotalSegmentCount}");
Console.WriteLine($"疑似未翻译数: {report.SuspectedUntranslatedCount}");
Console.WriteLine($"文本报告: {arguments.ReportTextPath}");
Console.WriteLine($"JSON 报告: {arguments.ReportJsonPath}");

return report.SuspectedUntranslatedCount > 0 ? 2 : 0;

internal static class Inspector
{
    public static InspectionReport Run(string pdfPath, string bilingualPath)
    {
        var segments = ReadSegments(bilingualPath);
        var findings = new List<InspectionFinding>();

        foreach (var segment in segments)
        {
            var reasons = DetectSuspiciousReasons(segment);
            if (reasons.Count == 0)
            {
                continue;
            }

            findings.Add(new InspectionFinding(
                segment.RowIndex,
                segment.Context,
                segment.Original,
                segment.Translation,
                reasons));
        }

        return new InspectionReport(
            Path.GetFullPath(pdfPath),
            Path.GetFullPath(bilingualPath),
            DateTimeOffset.Now,
            segments.Count,
            findings.Count,
            findings);
    }

    private static List<BilingualRow> ReadSegments(string bilingualPath)
    {
        var readablePath = CreateReadableCopy(bilingualPath);
        try
        {
            using var document = WordprocessingDocument.Open(readablePath, false);
            var body = document.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("DOCX 缺少正文。");
            var table = body.Elements<Table>().FirstOrDefault()
                ?? throw new InvalidOperationException("DOCX 中未找到双语表格。");

            var rows = table.Elements<TableRow>().ToList();
            var result = new List<BilingualRow>();

            for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                var cells = rows[rowIndex].Elements<TableCell>().ToList();
                if (cells.Count < 3)
                {
                    continue;
                }

                result.Add(new BilingualRow(
                    rowIndex,
                    NormalizeCellText(cells[0]),
                    NormalizeCellText(cells[1]),
                    NormalizeCellText(cells[2])));
            }

            return result;
        }
        finally
        {
            try
            {
                File.Delete(readablePath);
            }
            catch
            {
                // 临时文件删除失败不影响检查结果。
            }
        }
    }

    private static string CreateReadableCopy(string bilingualPath)
    {
        var tempDirectory = Path.Combine(Path.GetTempPath(), "PdfBilingualInspector");
        Directory.CreateDirectory(tempDirectory);
        var tempPath = Path.Combine(tempDirectory, $"{Guid.NewGuid():N}.docx");

        using var input = new FileStream(bilingualPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var output = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None);
        input.CopyTo(output);
        return tempPath;
    }

    private static string NormalizeCellText(TableCell cell)
    {
        var texts = cell.Descendants<Text>()
            .Select(x => x.Text)
            .Where(x => x is not null)
            .ToList();
        var joined = string.Join(string.Empty, texts);
        joined = joined.Replace("\r", "\n", StringComparison.Ordinal);
        joined = Regex.Replace(joined, @"\n{3,}", "\n\n");
        return joined.Trim();
    }

    private static List<string> DetectSuspiciousReasons(BilingualRow row)
    {
        var reasons = new List<string>();
        var original = row.Original.Trim();
        var translation = row.Translation.Trim();
        if (string.IsNullOrWhiteSpace(original))
        {
            return reasons;
        }

        if (string.IsNullOrWhiteSpace(translation))
        {
            reasons.Add("译文为空");
            return reasons;
        }

        if (IsTriviallyNonTranslatable(original))
        {
            return reasons;
        }

        var normalizedOriginal = NormalizeForComparison(original);
        var normalizedTranslation = NormalizeForComparison(translation);
        if (normalizedOriginal.Length > 0 &&
            string.Equals(normalizedOriginal, normalizedTranslation, StringComparison.OrdinalIgnoreCase))
        {
            reasons.Add("译文与原文相同");
        }

        if (ShouldFlagLongEnglishRun(original, translation))
        {
            reasons.Add("译文包含较长连续英文");
        }

        if (!ContainsCjk(translation) && EnglishLetterRatio(translation) > 0.45)
        {
            reasons.Add("译文几乎全是英文");
        }

        if (LooksLikeTrimmedContinuation(original, translation))
        {
            reasons.Add("译文疑似直接保留了英文续句");
        }

        return reasons;
    }

    private static string NormalizeForComparison(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            if (char.IsLetterOrDigit(ch))
            {
                builder.Append(char.ToLowerInvariant(ch));
            }
            else if (IsCjk(ch))
            {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }

    private static bool ContainsLongEnglishRun(string text) =>
        Regex.IsMatch(text, @"\b[A-Za-z][A-Za-z\-]{7,}\b");

    private static bool ShouldFlagLongEnglishRun(string original, string translation)
    {
        if (!ContainsLongEnglishRun(translation))
        {
            return false;
        }

        if (!ContainsCjk(translation))
        {
            return true;
        }

        var englishRatio = EnglishLetterRatio(translation);
        if (englishRatio > 0.38)
        {
            return true;
        }

        return NormalizeForComparison(translation).Contains(NormalizeForComparison(original), StringComparison.OrdinalIgnoreCase);
    }

    private static double EnglishLetterRatio(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return 0;
        }

        var letters = text.Count(ch => ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z');
        return letters / (double)Math.Max(1, text.Count(ch => !char.IsWhiteSpace(ch)));
    }

    private static bool ContainsCjk(string text) => text.Any(IsCjk);

    private static bool IsCjk(char ch) =>
        ch is >= '\u3400' and <= '\u4DBF' or
        >= '\u4E00' and <= '\u9FFF' or
        >= '\uF900' and <= '\uFAFF';

    private static bool LooksLikeTrimmedContinuation(string original, string translation)
    {
        if (translation.Length < 12)
        {
            return false;
        }

        var translationHead = translation[..Math.Min(36, translation.Length)];
        var originalHead = original[..Math.Min(36, original.Length)];
        return Regex.IsMatch(translationHead, @"^[A-Za-z][A-Za-z,\-\s()0-9]{10,}$") &&
               !ContainsCjk(translationHead) &&
               !string.Equals(
                   NormalizeForComparison(translationHead),
                   NormalizeForComparison(originalHead),
                   StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsTriviallyNonTranslatable(string text)
    {
        var compact = Regex.Replace(text, @"\s+", string.Empty);
        if (string.IsNullOrEmpty(compact))
        {
            return true;
        }

        if (!compact.Any(char.IsLetter) && compact.Any(char.IsDigit))
        {
            return true;
        }

        return false;
    }
}

internal static class ReportWriter
{
    public static string WriteText(InspectionReport report)
    {
        var builder = new StringBuilder();
        builder.AppendLine("PDF 双语导出未翻译检查报告");
        builder.AppendLine($"生成时间: {report.GeneratedAt:yyyy-MM-dd HH:mm:ss zzz}");
        builder.AppendLine($"PDF: {report.PdfPath}");
        builder.AppendLine($"双语 DOCX: {report.BilingualPath}");
        builder.AppendLine($"总片段数: {report.TotalSegmentCount}");
        builder.AppendLine($"疑似未翻译数: {report.SuspectedUntranslatedCount}");
        builder.AppendLine();

        if (report.Findings.Count == 0)
        {
            builder.AppendLine("未发现明显未翻译片段。");
            return builder.ToString();
        }

        foreach (var finding in report.Findings)
        {
            builder.AppendLine($"[行 {finding.RowIndex}] {finding.Context}");
            builder.AppendLine($"原因: {string.Join("；", finding.Reasons)}");
            builder.AppendLine($"原文: {TrimPreview(finding.Original)}");
            builder.AppendLine($"译文: {TrimPreview(finding.Translation)}");
            builder.AppendLine();
        }

        return builder.ToString();
    }

    private static string TrimPreview(string text)
    {
        var compact = Regex.Replace(text.Replace("\r", " ", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal), @"\s+", " ").Trim();
        return compact.Length <= 220 ? compact : compact[..220] + "...";
    }
}

internal sealed record CommandLineArguments(
    string PdfPath,
    string BilingualPath,
    string ReportTextPath,
    string ReportJsonPath,
    bool ShowHelp,
    string? ErrorMessage)
{
    public static CommandLineArguments Parse(string[] args)
    {
        try
        {
            return ParseCore(args);
        }
        catch (ArgumentException ex)
        {
            return new CommandLineArguments(string.Empty, string.Empty, string.Empty, string.Empty, false, ex.Message);
        }
    }

    private static CommandLineArguments ParseCore(string[] args)
    {
        if (args.Length == 0)
        {
            return new CommandLineArguments(string.Empty, string.Empty, string.Empty, string.Empty, true, null);
        }

        string? pdfPath = null;
        string? bilingualPath = null;
        string? reportBasePath = null;

        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--help":
                case "-h":
                    return new CommandLineArguments(string.Empty, string.Empty, string.Empty, string.Empty, true, null);
                case "--pdf":
                    pdfPath = ReadValue(args, ref i, "--pdf");
                    break;
                case "--bilingual":
                    bilingualPath = ReadValue(args, ref i, "--bilingual");
                    break;
                case "--report":
                    reportBasePath = ReadValue(args, ref i, "--report");
                    break;
                default:
                    return new CommandLineArguments(string.Empty, string.Empty, string.Empty, string.Empty, false, $"未知参数：{args[i]}");
            }
        }

        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            return new CommandLineArguments(string.Empty, string.Empty, string.Empty, string.Empty, false, "缺少必填参数 --pdf");
        }

        pdfPath = Path.GetFullPath(pdfPath);
        bilingualPath ??= Path.Combine(
            Path.GetDirectoryName(pdfPath)!,
            $"{Path.GetFileNameWithoutExtension(pdfPath)}.bilingual.docx");

        reportBasePath ??= Path.Combine(
            Path.GetDirectoryName(bilingualPath)!,
            $"{Path.GetFileNameWithoutExtension(bilingualPath)}.inspection");

        return new CommandLineArguments(
            pdfPath,
            Path.GetFullPath(bilingualPath),
            Path.GetFullPath(reportBasePath + ".txt"),
            Path.GetFullPath(reportBasePath + ".json"),
            false,
            null);
    }

    private static string ReadValue(string[] args, ref int index, string option)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException($"{option} 缺少值。");
        }

        index++;
        return args[index];
    }

    public static void PrintUsage()
    {
        Console.WriteLine("用法:");
        Console.WriteLine("  dotnet run --project tools/PdfBilingualInspector -- --pdf <原始PDF路径> [--bilingual <双语DOCX路径>] [--report <报告输出前缀>]");
        Console.WriteLine();
        Console.WriteLine("示例:");
        Console.WriteLine(@"  dotnet run --project tools/PdfBilingualInspector -- --pdf ""C:\docs\paper.pdf""");
    }
}

internal sealed record BilingualRow(int RowIndex, string Context, string Original, string Translation);

internal sealed record InspectionFinding(
    int RowIndex,
    string Context,
    string Original,
    string Translation,
    IReadOnlyList<string> Reasons);

internal sealed record InspectionReport(
    string PdfPath,
    string BilingualPath,
    DateTimeOffset GeneratedAt,
    int TotalSegmentCount,
    int SuspectedUntranslatedCount,
    IReadOnlyList<InspectionFinding> Findings);
