using System.Collections.Concurrent;
using System.IO;
using PdfSharp.Fonts;

namespace TranslatorApp.Infrastructure;

public static class PdfSharpFontResolver
{
    private static int _initialized;
    public const string DefaultFontFamily = "DengXian";

    public static void Initialize()
    {
        if (Interlocked.Exchange(ref _initialized, 1) == 1)
        {
            return;
        }

        GlobalFontSettings.FontResolver = new WindowsCjkFontResolver();
    }

    private sealed class WindowsCjkFontResolver : IFontResolver
    {
        private static readonly string FontsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

        // Sans-serif / 无衬线字体候选列表（常规、粗体、斜体、粗斜体）
        private static readonly string[] PreferredSansFiles =
        [
            "Deng.ttf",
            "msyh.ttc",
            "simhei.ttf",
            "arial.ttf"
        ];

        private static readonly string[] PreferredSansBoldFiles =
        [
            "Dengb.ttf",
            "msyhbd.ttc",
            "simhei.ttf",
            "arialbd.ttf",
            "Deng.ttf"
        ];

        private static readonly string[] PreferredSansItalicFiles =
        [
            "Dengl.ttf",
            "Deng.ttf",
            "ariali.ttf",
            "msyh.ttc",
            "simhei.ttf"
        ];

        private static readonly string[] PreferredSansBoldItalicFiles =
        [
            "Dengb.ttf",
            "msyhbd.ttc",
            "arialbi.ttf",
            "simhei.ttf"
        ];

        // Serif / 衬线字体候选列表
        private static readonly string[] PreferredSerifFiles =
        [
            "times.ttf",
            "simsun.ttc",
            "Deng.ttf",
            "arial.ttf"
        ];

        private static readonly string[] PreferredSerifBoldFiles =
        [
            "timesbd.ttf",
            "simsun.ttc",
            "Dengb.ttf",
            "arialbd.ttf"
        ];

        private static readonly string[] PreferredSerifItalicFiles =
        [
            "timesi.ttf",
            "simsun.ttc",
            "Deng.ttf",
            "ariali.ttf"
        ];

        private static readonly string[] PreferredSerifBoldItalicFiles =
        [
            "timesbi.ttf",
            "simsun.ttc",
            "Dengb.ttf",
            "arialbi.ttf"
        ];

        // Monospace / 等宽字体候选列表
        private static readonly string[] PreferredMonospaceFiles =
        [
            "consola.ttf",
            "cour.ttf",
            "simhei.ttf",
            "Deng.ttf"
        ];

        private static readonly string[] PreferredMonospaceBoldFiles =
        [
            "consolab.ttf",
            "courbd.ttf",
            "simhei.ttf",
            "Dengb.ttf"
        ];

        private static readonly string[] PreferredMonospaceItalicFiles =
        [
            "consolai.ttf",
            "couri.ttf",
            "consola.ttf",
            "Deng.ttf"
        ];

        private static readonly string[] PreferredMonospaceBoldItalicFiles =
        [
            "consolaz.ttf",
            "courbi.ttf",
            "consolab.ttf",
            "Dengb.ttf"
        ];

        private static readonly ConcurrentDictionary<string, byte[]> FontBytesCache = new(StringComparer.OrdinalIgnoreCase);

        public byte[]? GetFont(string faceName)
        {
            if (string.IsNullOrWhiteSpace(faceName) || !File.Exists(faceName))
            {
                return null;
            }

            return FontBytesCache.GetOrAdd(faceName, static key => File.ReadAllBytes(key));
        }

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            var candidates = GetCandidateFiles(familyName, isBold, isItalic);
            var path = candidates
                .Select(file => Path.Combine(FontsDirectory, file))
                .FirstOrDefault(File.Exists);

            if (path is null)
            {
                // 第一次回退：尝试不带样式的基础字体
                if (isBold || isItalic)
                {
                    var baseCandidates = GetCandidateFiles(familyName, false, false);
                    path = baseCandidates
                        .Select(file => Path.Combine(FontsDirectory, file))
                        .FirstOrDefault(File.Exists);
                }
            }

            if (path is null)
            {
                // 第二次回退：尝试默认字体族
                var defaultCandidates = GetCandidateFiles(DefaultFontFamily, isBold, isItalic);
                path = defaultCandidates
                    .Select(file => Path.Combine(FontsDirectory, file))
                    .FirstOrDefault(File.Exists);
            }

            if (path is null)
            {
                // 最终回退：扫描所有候选列表
                path = FindAnyInstalledFallback();
            }

            if (path is null)
            {
                throw new InvalidOperationException(
                    $"未找到可用字体文件。请求字体：{familyName}，样式：bold={isBold}, italic={isItalic}。");
            }

            return new FontResolverInfo(path);
        }

        private static IEnumerable<string> GetCandidateFiles(string familyName, bool isBold, bool isItalic)
        {
            var normalized = NormalizeFamilyName(familyName);

            // 等宽字体族
            if (normalized.Contains("courier") || normalized.Contains("consolas") ||
                normalized.Contains("mono") || normalized.Contains("menlo") ||
                normalized.Contains("fira code") || normalized.Contains("source code") ||
                normalized.Contains("jetbrains") || normalized.Contains("inconsolata") ||
                normalized.Contains("hack"))
            {
                return (isBold, isItalic) switch
                {
                    (true, true) => PreferredMonospaceBoldItalicFiles,
                    (true, false) => PreferredMonospaceBoldFiles,
                    (false, true) => PreferredMonospaceItalicFiles,
                    _ => PreferredMonospaceFiles
                };
            }

            // 衬线字体族
            if (normalized.Contains("times") || normalized.Contains("serif") ||
                normalized.Contains("song") || normalized.Contains("simsun") ||
                normalized.Contains("songti") || normalized.Contains("宋") ||
                normalized.Contains("georgia") || normalized.Contains("cambria") ||
                normalized.Contains("palatino") || normalized.Contains("garamond") ||
                normalized.Contains("book antiqua") || normalized.Contains("mincho") ||
                normalized.Contains("明朝") || normalized.Contains("noto serif"))
            {
                return (isBold, isItalic) switch
                {
                    (true, true) => PreferredSerifBoldItalicFiles,
                    (true, false) => PreferredSerifBoldFiles,
                    (false, true) => PreferredSerifItalicFiles,
                    _ => PreferredSerifFiles
                };
            }

            // Arial / Helvetica 精确匹配
            if (normalized.Contains("arial") || normalized.Contains("helvetica"))
            {
                return (isBold, isItalic) switch
                {
                    (true, true) => ["arialbi.ttf", ..PreferredSansBoldItalicFiles],
                    (true, false) => ["arialbd.ttf", ..PreferredSansBoldFiles],
                    (false, true) => ["ariali.ttf", ..PreferredSansItalicFiles],
                    _ => ["arial.ttf", ..PreferredSansFiles]
                };
            }

            // Calibri / Segoe / Verdana / Tahoma 映射到 Arial 相近字体
            if (normalized.Contains("calibri") || normalized.Contains("segoe") ||
                normalized.Contains("verdana") || normalized.Contains("tahoma") ||
                normalized.Contains("trebuchet"))
            {
                return (isBold, isItalic) switch
                {
                    (true, true) => ["arialbi.ttf", ..PreferredSansBoldItalicFiles],
                    (true, false) => ["arialbd.ttf", ..PreferredSansBoldFiles],
                    (false, true) => ["ariali.ttf", ..PreferredSansItalicFiles],
                    _ => ["arial.ttf", ..PreferredSansFiles]
                };
            }

            // 微软雅黑
            if (normalized.Contains("yahei") || normalized.Contains("微软雅黑"))
            {
                return (isBold, isItalic) switch
                {
                    (true, true) => ["msyhbd.ttc", ..PreferredSansBoldItalicFiles],
                    (true, false) => ["msyhbd.ttc", ..PreferredSansBoldFiles],
                    (false, true) => ["msyhl.ttc", "msyh.ttc", ..PreferredSansItalicFiles],
                    _ => ["msyh.ttc", ..PreferredSansFiles]
                };
            }

            // 其他无衬线 / CJK Sans 字体
            if (normalized.Contains("noto") || normalized.Contains("deng") ||
                normalized.Contains("hei") || normalized.Contains("sans") ||
                normalized.Contains("黑") || normalized.Contains("等线") ||
                normalized.Contains("gothic") || normalized.Contains("ゴシック"))
            {
                return (isBold, isItalic) switch
                {
                    (true, true) => PreferredSansBoldItalicFiles,
                    (true, false) => PreferredSansBoldFiles,
                    (false, true) => PreferredSansItalicFiles,
                    _ => PreferredSansFiles
                };
            }

            // 楷体
            if (normalized.Contains("kaiti") || normalized.Contains("楷"))
            {
                // 楷体没有粗体/斜体变体，回退到无衬线
                return isBold ? PreferredSansBoldFiles : PreferredSansFiles;
            }

            // 仿宋
            if (normalized.Contains("fangsong") || normalized.Contains("仿宋"))
            {
                return isBold ? PreferredSerifBoldFiles : PreferredSerifFiles;
            }

            // 默认：按样式选择无衬线字体
            return (isBold, isItalic) switch
            {
                (true, true) => PreferredSansBoldItalicFiles,
                (true, false) => PreferredSansBoldFiles,
                (false, true) => PreferredSansItalicFiles,
                _ => PreferredSansFiles
            };
        }

        private static string NormalizeFamilyName(string familyName) =>
            string.IsNullOrWhiteSpace(familyName)
                ? DefaultFontFamily.ToLowerInvariant()
                : familyName
                    .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Select(candidate => candidate.Trim().Trim('"', '\''))
                    .FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate))?
                    .ToLowerInvariant() ?? DefaultFontFamily.ToLowerInvariant();

        private static string? FindAnyInstalledFallback()
        {
            var orderedCandidates = PreferredSansFiles
                .Concat(PreferredSansBoldFiles)
                .Concat(PreferredSansItalicFiles)
                .Concat(PreferredSansBoldItalicFiles)
                .Concat(PreferredSerifFiles)
                .Concat(PreferredSerifBoldFiles)
                .Concat(PreferredMonospaceFiles)
                .Concat(PreferredMonospaceBoldFiles)
                .Distinct(StringComparer.OrdinalIgnoreCase);

            return orderedCandidates
                .Select(file => Path.Combine(FontsDirectory, file))
                .FirstOrDefault(File.Exists);
        }
    }
}
