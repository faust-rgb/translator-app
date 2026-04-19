using System.Collections.Concurrent;
using TranslatorApp.Configuration;
using TranslatorApp.Models;
using TranslatorApp.Services.Ai;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;

namespace TranslatorApp.Services;

public sealed class TextTranslationService(
    IAiTranslationClientFactory clientFactory,
    IGlossaryService glossaryService,
    IAppLogService logService,
    ITranslationRequestThrottle translationRequestThrottle) : ITextTranslationService
{
    private readonly ConcurrentDictionary<string, Lazy<Task<string>>> _translationCache = new(StringComparer.Ordinal);
    private static readonly Regex UrlRegex = new(@"https?://\S+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EmailRegex = new(@"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DoiRegex = new(@"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ArxivRegex = new(@"\barXiv:\d{4}\.\d{4,5}(?:v\d+)?\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex LatexCommandRegex = new(@"\\[A-Za-z]+(?:\{[^{}]*\})?", RegexOptions.Compiled);
    private static readonly Regex CitationBracketRegex = new(@"\[(?:\d{1,3}(?:\s*[-,–]\s*\d{1,3}){0,8}|[A-Za-z][A-Za-z0-9_.-]{0,20})\]", RegexOptions.Compiled);
    private static readonly Regex SubfigureLabelRegex = new(@"(?<!\w)\([a-z]\)(?!\w)", RegexOptions.Compiled);
    private static readonly Regex SectionNumberPrefixRegex = new(@"(?m)^(?:\s*)(?:第[一二三四五六七八九十百千万0-9]+[章节条款部分]|(?:Chapter|Section|Article)\s+\d+[A-Za-z0-9.\-]*|(?:\d{1,3}|[A-Za-z]|[IVXLC]+)(?:\.\d{1,3}){0,5}[\.\)、\)]|[\(（](?:\d{1,3}|[A-Za-z]|[ivxlc]+)[\)）]|[A-Za-z]\)|[•\-·▪■◆◦])\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex InlineClauseNumberRegex = new(@"(?<!\w)(?:第[一二三四五六七八九十百千万0-9]+[章节条款部分]|(?:\d{1,3})(?:\.\d{1,3}){1,5}|(?:Article|Section)\s+\d+[A-Za-z0-9.\-]*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public async Task<string> TranslateAsync(
        string text,
        string contextHint,
        string additionalRequirements,
        AppSettings settings,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        var glossaryEntries = await glossaryService.LoadAsync(settings.Translation.GlossaryPath, cancellationToken);
        var protectedText = ProtectSpecialSegments(text);
        var composedRequirements = ComposeAdditionalRequirements(additionalRequirements, protectedText);
        var request = new TranslationRequest
        {
            Text = protectedText.Text,
            SourceLanguage = settings.Translation.SourceLanguage,
            TargetLanguage = settings.Translation.TargetLanguage,
            ContextHint = contextHint,
            PromptTemplate = settings.Translation.PromptTemplate,
            GlossaryPrompt = glossaryService.BuildPromptSection(text, glossaryEntries),
            AdditionalRequirements = composedRequirements,
            EnableStreaming = settings.Translation.EnableStreaming && onPartialResponse is not null,
            OnPartialResponse = onPartialResponse
        };

        var cacheKey = BuildCacheKey(settings, request);
        var lazyTranslation = _translationCache.GetOrAdd(
            cacheKey,
            _ => new Lazy<Task<string>>(
                () => TranslateCoreAsync(text, protectedText, request, settings, cancellationToken),
                LazyThreadSafetyMode.ExecutionAndPublication));

        try
        {
            return await lazyTranslation.Value;
        }
        catch
        {
            _translationCache.TryRemove(cacheKey, out _);
            throw;
        }
    }

    private async Task<string> TranslateCoreAsync(
        string originalText,
        ProtectedText protectedText,
        TranslationRequest request,
        AppSettings settings,
        CancellationToken cancellationToken)
    {
        var serverErrorRetryCount = ResolveServerErrorRetryCount(settings.Translation);
        var timeoutRetryCount = Math.Max(0, settings.Translation.TimeoutRetryCount);
        var exceptions = new List<ExceptionDispatchInfo>();
        var serverErrorRetriesUsed = 0;
        var timeoutRetriesUsed = 0;

        while (true)
        {
            try
            {
                var client = clientFactory.Create(settings.Ai);
                await using var _ = await translationRequestThrottle.AcquireAsync(
                    settings.Translation.MaxGlobalTranslationRequests,
                    cancellationToken);
                var rawTranslated = await client.TranslateAsync(request, cancellationToken);
                var translated = protectedText.Restore(rawTranslated);
                if (ShouldRetryForQuality(originalText, translated, rawTranslated, request.AdditionalRequirements, protectedText))
                {
                    client = clientFactory.Create(settings.Ai);
                    var retryRequest = new TranslationRequest
                    {
                        Text = request.Text,
                        SourceLanguage = request.SourceLanguage,
                        TargetLanguage = request.TargetLanguage,
                        ContextHint = request.ContextHint,
                        PromptTemplate = request.PromptTemplate,
                        GlossaryPrompt = request.GlossaryPrompt,
                        EnableStreaming = false,
                        OnPartialResponse = null,
                        AdditionalRequirements = BuildRetryRequirements(request.AdditionalRequirements, protectedText)
                    };
                    rawTranslated = await client.TranslateAsync(retryRequest, cancellationToken);
                    translated = protectedText.Restore(rawTranslated);
                }

                return translated;
            }
            catch (Exception ex) when (ShouldRetryTimeout(ex, cancellationToken) && timeoutRetriesUsed < timeoutRetryCount)
            {
                exceptions.Add(ExceptionDispatchInfo.Capture(ex));
                timeoutRetriesUsed++;
                logService.Error($"翻译请求超时，第 {timeoutRetriesUsed} 次超时重试前等待：{ex.Message}");
                await Task.Delay(GetRetryDelay(timeoutRetriesUsed), cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException && serverErrorRetriesUsed < serverErrorRetryCount)
            {
                exceptions.Add(ExceptionDispatchInfo.Capture(ex));
                serverErrorRetriesUsed++;
                logService.Error($"翻译请求失败，第 {serverErrorRetriesUsed} 次异常重试前等待：{ex.Message}");
                await Task.Delay(GetRetryDelay(serverErrorRetriesUsed), cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                exceptions.Add(ExceptionDispatchInfo.Capture(ex));
                break;
            }
        }

        if (exceptions.Count == 0)
        {
            throw new InvalidOperationException("翻译失败。");
        }

        if (exceptions.Count == 1)
        {
            exceptions[0].Throw();
        }

        throw new AggregateException(
            $"翻译在 {exceptions.Count} 次尝试后仍失败。",
            exceptions.Select(x => x.SourceException));
    }

    private static string BuildCacheKey(AppSettings settings, TranslationRequest request) =>
        string.Join(
            "\u001F",
            settings.Ai.ProviderType,
            settings.Ai.BaseUrl ?? string.Empty,
            settings.Ai.Model ?? string.Empty,
            request.SourceLanguage ?? string.Empty,
            request.TargetLanguage ?? string.Empty,
            request.Text ?? string.Empty,
            request.GlossaryPrompt ?? string.Empty,
            request.AdditionalRequirements ?? string.Empty,
            request.PromptTemplate ?? string.Empty);

    private static bool ShouldRetryForQuality(
        string original,
        string translated,
        string rawTranslated,
        string additionalRequirements,
        ProtectedText protectedText)
    {
        if (protectedText.HasTokens && protectedText.CountMissingPlaceholders(rawTranslated) > 0)
        {
            return true;
        }

        if (ShouldRetryForUntranslatedFragment(original, translated))
        {
            return true;
        }

        if (IsPdfTableLike(additionalRequirements))
        {
            if (LooksSuspiciouslyUntranslated(translated))
            {
                return true;
            }

            if (LooksLikeTruncatedStructuredTranslation(original, translated))
            {
                return true;
            }
        }

        if (IsPdfCaptionLike(additionalRequirements))
        {
            if (LooksSuspiciouslyUntranslated(translated))
            {
                return true;
            }

            if (protectedText.HasTokens && !protectedText.AllProtectedContentRestored(translated))
            {
                return true;
            }
        }

        if (IsWordListLike(additionalRequirements) && LooksLikeWordListStructureBroken(original, translated))
        {
            return true;
        }

        if (LooksLikeNumberingStructureBroken(original, translated))
        {
            return true;
        }

        if (IsWordTableCellLike(additionalRequirements))
        {
            if (LooksSuspiciouslyUntranslated(translated))
            {
                return true;
            }

            if (LooksLikeTruncatedStructuredTranslation(original, translated))
            {
                return true;
            }

            if (LooksLikeWordTableCellOverExpanded(original, translated))
            {
                return true;
            }
        }

        if (IsWordBoundarySensitive(additionalRequirements))
        {
            if (LooksSuspiciouslyUntranslated(translated))
            {
                return true;
            }

            if (protectedText.HasTokens && !protectedText.AllProtectedContentRestored(translated))
            {
                return true;
            }

            if (LooksLikeBoundarySensitiveStructureBroken(original, translated, additionalRequirements))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldRetryForUntranslatedFragment(string original, string translated)
    {
        if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(translated))
        {
            return false;
        }

        var originalTrimmed = original.Trim();
        var translatedTrimmed = translated.Trim();
        if (!LooksLikeRiskyPdfFragment(originalTrimmed))
        {
            return false;
        }

        if (NormalizeForComparison(originalTrimmed) == NormalizeForComparison(translatedTrimmed))
        {
            return true;
        }

        return EnglishLetterRatio(translatedTrimmed) > 0.55 && !ContainsCjk(translatedTrimmed);
    }

    private static bool LooksLikeRiskyPdfFragment(string text)
    {
        if (text.Length < 20)
        {
            return false;
        }

        var trimmedStart = text.TrimStart();
        return char.IsLower(trimmedStart[0]) ||
               EndsWithHyphenLikeBreak(text) ||
               Regex.IsMatch(text, @"\b(?:in|of|for|to|with|from|by|and|or|but)\s*$", RegexOptions.IgnoreCase);
    }

    private static bool EndsWithHyphenLikeBreak(string text)
    {
        var end = text.TrimEnd();
        return end.EndsWith('-') ||
               end.EndsWith('‐') ||
               end.EndsWith('‑') ||
               end.EndsWith('‒') ||
               end.EndsWith('–');
    }

    private static string NormalizeForComparison(string text)
    {
        var chars = text.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant);
        return new string(chars.ToArray());
    }

    private static double EnglishLetterRatio(string text)
    {
        var nonWhitespace = text.Count(ch => !char.IsWhiteSpace(ch));
        if (nonWhitespace == 0)
        {
            return 0;
        }

        var letters = text.Count(ch => ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z');
        return letters / (double)nonWhitespace;
    }

    private static bool ContainsCjk(string text) =>
        text.Any(ch => ch is >= '\u3400' and <= '\u4DBF' or >= '\u4E00' and <= '\u9FFF' or >= '\uF900' and <= '\uFAFF');

    private static int ResolveServerErrorRetryCount(TranslationSettings settings) =>
        settings.ServerErrorRetryCount == 2 && settings.RetryCount != 2
            ? Math.Max(0, settings.RetryCount)
            : Math.Max(0, settings.ServerErrorRetryCount);

    private static bool ShouldRetryTimeout(Exception ex, CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            return false;
        }

        return ex is TimeoutException
               || ex is TaskCanceledException
               || ex.InnerException is TimeoutException;
    }

    private static string ComposeAdditionalRequirements(string additionalRequirements, ProtectedText protectedText)
    {
        if (!protectedText.HasTokens)
        {
            return additionalRequirements ?? string.Empty;
        }

        var protectionNote = "必须原样保留文中出现的 @@KEEP_xxxx@@ 占位符，不要改写、拆分或删除；输出时也保留这些占位符。";
        return string.IsNullOrWhiteSpace(additionalRequirements)
            ? protectionNote
            : $"{additionalRequirements.Trim()}\n{protectionNote}";
    }

    private static string BuildRetryRequirements(string additionalRequirements, ProtectedText protectedText)
    {
        var retryNote =
            "上一次返回仍存在漏译或结构破坏风险。请完整翻译自然语言内容，但必须严格保留公式、变量、编号、单位、引用、占位符以及原有结构。" +
            "不要照抄大段英文，不要把表格/图注多列内容合并成解释性整段。";
        var baseRequirements = string.IsNullOrWhiteSpace(additionalRequirements)
            ? retryNote
            : $"{additionalRequirements.Trim()}\n{retryNote}";

        return protectedText.HasTokens
            ? ComposeAdditionalRequirements(baseRequirements, protectedText)
            : baseRequirements;
    }

    private static bool IsPdfTableLike(string additionalRequirements) =>
        additionalRequirements.Contains("类型：表格", StringComparison.Ordinal) ||
        additionalRequirements.Contains("类型：列表", StringComparison.Ordinal);

    private static bool IsPdfCaptionLike(string additionalRequirements) =>
        additionalRequirements.Contains("类型：图注", StringComparison.Ordinal) ||
        additionalRequirements.Contains("类型：标题", StringComparison.Ordinal) ||
        additionalRequirements.Contains("类型：脚注", StringComparison.Ordinal);

    private static bool IsWordListLike(string additionalRequirements) =>
        additionalRequirements.Contains("类型：列表", StringComparison.Ordinal);

    private static bool IsWordTableCellLike(string additionalRequirements) =>
        additionalRequirements.Contains("类型：表格单元格", StringComparison.Ordinal);

    private static bool IsWordBoundarySensitive(string additionalRequirements) =>
        additionalRequirements.Contains("超链接显示文本", StringComparison.Ordinal) ||
        additionalRequirements.Contains("上标或下标", StringComparison.Ordinal) ||
        additionalRequirements.Contains("字段结果", StringComparison.Ordinal) ||
        additionalRequirements.Contains("文本框", StringComparison.Ordinal) ||
        additionalRequirements.Contains("类型：脚注/尾注", StringComparison.Ordinal);

    private static bool LooksSuspiciouslyUntranslated(string translated)
    {
        if (string.IsNullOrWhiteSpace(translated))
        {
            return false;
        }

        return EnglishLetterRatio(translated.Trim()) > 0.55 && !ContainsCjk(translated);
    }

    private static bool LooksLikeTruncatedStructuredTranslation(string original, string translated)
    {
        var originalTrimmed = original.Trim();
        var translatedTrimmed = translated.Trim();
        if (originalTrimmed.Length < 32 || translatedTrimmed.Length == 0)
        {
            return false;
        }

        var originalTokens = Regex.Matches(originalTrimmed, @"\S+").Count;
        var translatedTokens = Regex.Matches(translatedTrimmed, @"\S+").Count;
        return translatedTokens <= Math.Max(1, originalTokens / 4) &&
               translatedTrimmed.Length <= Math.Max(12, originalTrimmed.Length / 4);
    }

    private static bool LooksLikeWordListStructureBroken(string original, string translated)
    {
        var originalLines = SplitNonEmptyLines(original);
        var translatedLines = SplitNonEmptyLines(translated);
        if (originalLines.Count < 2)
        {
            return false;
        }

        if (translatedLines.Count < Math.Max(1, originalLines.Count - 1))
        {
            return true;
        }

        var originalBulletLike = originalLines.Count(IsBulletLikeLine);
        if (originalBulletLike == 0)
        {
            return false;
        }

        var translatedBulletLike = translatedLines.Count(IsBulletLikeLine);
        return translatedBulletLike < Math.Max(1, originalBulletLike - 1);
    }

    private static bool LooksLikeWordTableCellOverExpanded(string original, string translated)
    {
        var originalTrimmed = original.Trim();
        var translatedTrimmed = translated.Trim();
        if (string.IsNullOrWhiteSpace(originalTrimmed) || string.IsNullOrWhiteSpace(translatedTrimmed))
        {
            return false;
        }

        var originalLineCount = originalTrimmed.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length;
        var translatedLineCount = translatedTrimmed.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length;
        if (originalLineCount > 1 && translatedLineCount == 1 && originalTrimmed.Length <= 80)
        {
            return true;
        }

        var originalTokenCount = Regex.Matches(originalTrimmed, @"\S+").Count;
        var translatedTokenCount = Regex.Matches(translatedTrimmed, @"\S+").Count;
        var originalLooksCompact = originalTokenCount <= 10 && originalTrimmed.Length <= 80;
        return originalLooksCompact &&
               translatedTokenCount > Math.Max(14, originalTokenCount * 2) &&
               translatedTrimmed.Length > Math.Max(80, originalTrimmed.Length * 2.4);
    }

    private static bool LooksLikeBoundarySensitiveStructureBroken(string original, string translated, string additionalRequirements)
    {
        if (additionalRequirements.Contains("上标或下标", StringComparison.Ordinal) &&
            CountSuperscriptLikeMarkers(original) > CountSuperscriptLikeMarkers(translated))
        {
            return true;
        }

        if (additionalRequirements.Contains("文本框", StringComparison.Ordinal))
        {
            var originalLooksShort = original.Trim().Length <= 80;
            var translatedLooksLong = translated.Trim().Length > Math.Max(100, original.Trim().Length * 2.5);
            if (originalLooksShort && translatedLooksLong)
            {
                return true;
            }
        }

        return false;
    }

    private static bool LooksLikeNumberingStructureBroken(string original, string translated)
    {
        var originalPrefixes = ExtractLinePrefixes(original);
        if (originalPrefixes.Count == 0)
        {
            return false;
        }

        var translatedPrefixes = ExtractLinePrefixes(translated);
        if (translatedPrefixes.Count < Math.Max(1, originalPrefixes.Count - 1))
        {
            return true;
        }

        var originalInlineClauses = InlineClauseNumberRegex.Matches(original).Count;
        if (originalInlineClauses == 0)
        {
            return false;
        }

        var translatedInlineClauses = InlineClauseNumberRegex.Matches(translated).Count;
        return translatedInlineClauses < Math.Max(1, originalInlineClauses - 1);
    }

    private static int CountSuperscriptLikeMarkers(string text) =>
        Regex.Matches(text, @"(?:(?<=\d)[A-Za-z](?=\b)|(?<=\b)[A-Za-z](?=\d)|\[\d+\])").Count;

    private static List<string> SplitNonEmptyLines(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => line.Length > 0)
            .ToList();

    private static bool IsBulletLikeLine(string line) =>
        Regex.IsMatch(line, @"^\s*(?:[\u2022\u25E6\u25AA\u25A0\u2043\-·▪■★]|(?:\d{1,3}|[A-Za-z])[\.\)、\)])\s*");

    private static List<string> ExtractLinePrefixes(string text) =>
        SectionNumberPrefixRegex.Matches(text)
            .Select(match => match.Value.Trim())
            .Where(value => value.Length > 0)
            .ToList();

    private static ProtectedText ProtectSpecialSegments(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return new ProtectedText(text ?? string.Empty, []);
        }

        var patterns = new[]
        {
            UrlRegex,
            EmailRegex,
            DoiRegex,
            ArxivRegex,
            LatexCommandRegex,
            CitationBracketRegex,
            SubfigureLabelRegex,
            SectionNumberPrefixRegex,
            InlineClauseNumberRegex
        };

        var candidates = new List<(int Index, int Length, string Value)>();
        foreach (var pattern in patterns)
        {
            foreach (Match match in pattern.Matches(text))
            {
                if (!match.Success || string.IsNullOrWhiteSpace(match.Value))
                {
                    continue;
                }

                candidates.Add((match.Index, match.Length, match.Value));
            }
        }

        if (candidates.Count == 0)
        {
            return new ProtectedText(text, []);
        }

        var selected = new List<(int Index, int Length, string Value)>();
        foreach (var candidate in candidates
                     .OrderBy(x => x.Index)
                     .ThenByDescending(x => x.Length))
        {
            var overlaps = selected.Any(existing =>
                candidate.Index < existing.Index + existing.Length &&
                existing.Index < candidate.Index + candidate.Length);
            if (!overlaps)
            {
                selected.Add(candidate);
            }
        }

        if (selected.Count == 0)
        {
            return new ProtectedText(text, []);
        }

        var builder = new System.Text.StringBuilder(text);
        var tokens = new List<ProtectedToken>(selected.Count);
        for (var i = selected.Count - 1; i >= 0; i--)
        {
            var item = selected[i];
            var placeholder = $"@@KEEP_{i + 1:0000}@@";
            builder.Remove(item.Index, item.Length);
            builder.Insert(item.Index, placeholder);
            tokens.Add(new ProtectedToken(placeholder, item.Value));
        }

        tokens.Reverse();
        return new ProtectedText(builder.ToString(), tokens);
    }

    private static TimeSpan GetRetryDelay(int retryIndex) =>
        TimeSpan.FromSeconds(Math.Min(8, 2 * retryIndex));

    private sealed record ProtectedToken(string Placeholder, string Original);

    private sealed record ProtectedText(string Text, IReadOnlyList<ProtectedToken> Tokens)
    {
        public bool HasTokens => Tokens.Count > 0;

        public int CountMissingPlaceholders(string translated) =>
            Tokens.Count(token => !translated.Contains(token.Placeholder, StringComparison.Ordinal));

        public bool AllProtectedContentRestored(string translated) =>
            Tokens.All(token => translated.Contains(token.Original, StringComparison.Ordinal));

        public string Restore(string translated)
        {
            translated ??= string.Empty;
            foreach (var token in Tokens)
            {
                translated = translated.Replace(token.Placeholder, token.Original, StringComparison.Ordinal);
            }

            return translated;
        }
    }
}
