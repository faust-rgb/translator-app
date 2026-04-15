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

        private static readonly string[] PreferredSansFiles =
        [
            "Deng.ttf",
            "simhei.ttf",
            "arial.ttf"
        ];

        private static readonly string[] PreferredSansBoldFiles =
        [
            "Dengb.ttf",
            "simhei.ttf",
            "arialbd.ttf",
            "Deng.ttf"
        ];

        private static readonly string[] PreferredSansItalicFiles =
        [
            "Deng.ttf",
            "ariali.ttf",
            "simhei.ttf",
            "arial.ttf"
        ];

        private static readonly string[] PreferredSerifFiles =
        [
            "arial.ttf",
            "Deng.ttf",
            "simhei.ttf"
        ];

        private static readonly string[] PreferredMonospaceFiles =
        [
            "consola.ttf",
            "cour.ttf",
            "simhei.ttf"
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
                path = FindAnyInstalledFallback();
            }

            return path is null ? null : new FontResolverInfo(path);
        }

        private static IEnumerable<string> GetCandidateFiles(string familyName, bool isBold, bool isItalic)
        {
            var normalized = NormalizeFamilyName(familyName);

            if (normalized.Contains("courier") || normalized.Contains("consolas") || normalized.Contains("mono"))
            {
                return PreferredMonospaceFiles;
            }

            if (normalized.Contains("times") || normalized.Contains("serif") || normalized.Contains("song"))
            {
                return PreferredSerifFiles;
            }

            if (normalized.Contains("arial"))
            {
                return isBold
                    ? ["arialbd.ttf", "arial.ttf", ..PreferredSansBoldFiles]
                    : isItalic
                        ? ["ariali.ttf", "arial.ttf", ..PreferredSansItalicFiles]
                        : ["arial.ttf", ..PreferredSansFiles];
            }

            if (normalized.Contains("yahei") || normalized.Contains("wei ruan ya hei") || normalized.Contains("微软雅黑"))
            {
                return isBold
                    ? PreferredSansBoldFiles
                    : isItalic
                        ? PreferredSansItalicFiles
                        : PreferredSansFiles;
            }

            if (normalized.Contains("noto") || normalized.Contains("deng") || normalized.Contains("hei") || normalized.Contains("sans"))
            {
                return isBold
                    ? PreferredSansBoldFiles
                    : PreferredSansFiles;
            }

            if (normalized.Contains("simsun") || normalized.Contains("songti"))
            {
                return PreferredSerifFiles;
            }

            return isBold
                ? PreferredSansBoldFiles
                : isItalic
                    ? PreferredSansItalicFiles
                    : PreferredSansFiles;
        }

        private static string NormalizeFamilyName(string familyName) =>
            string.IsNullOrWhiteSpace(familyName)
                ? DefaultFontFamily.ToLowerInvariant()
                : familyName.Trim().ToLowerInvariant();

        private static string? FindAnyInstalledFallback()
        {
            var orderedCandidates = PreferredSansFiles
                .Concat(PreferredSansBoldFiles)
                .Concat(PreferredSerifFiles)
                .Concat(PreferredMonospaceFiles)
                .Distinct(StringComparer.OrdinalIgnoreCase);

            return orderedCandidates
                .Select(file => Path.Combine(FontsDirectory, file))
                .FirstOrDefault(File.Exists);
        }
    }
}
