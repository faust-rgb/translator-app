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
    public async Task<string> TranslateAsync(
        string text,
        string contextHint,
        AppSettings settings,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        var glossaryEntries = await glossaryService.LoadAsync(settings.Translation.GlossaryPath, cancellationToken);
        var request = new TranslationRequest
        {
            Text = text,
            SourceLanguage = settings.Translation.SourceLanguage,
            TargetLanguage = settings.Translation.TargetLanguage,
            ContextHint = contextHint,
            PromptTemplate = settings.Translation.PromptTemplate,
            GlossaryPrompt = glossaryService.BuildPromptSection(text, glossaryEntries),
            EnableStreaming = settings.Translation.EnableStreaming && onPartialResponse is not null,
            OnPartialResponse = onPartialResponse
        };

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
                var translated = await client.TranslateAsync(request, cancellationToken);
                if (ShouldRetryForUntranslatedFragment(text, translated))
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
                        AdditionalRequirements =
                            "当前片段上一次返回仍保留了大段英文。请把这段英文正文完整译成中文。" +
                            "即使它是句中残片、断词续句或以下一个小写单词开头，也不要原样照抄英文。" +
                            "只有作者名、机构缩写、邮箱、公式符号等确实不应翻译的内容才可以保留。"
                    };
                    translated = await client.TranslateAsync(retryRequest, cancellationToken);
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

    private static TimeSpan GetRetryDelay(int retryIndex) =>
        TimeSpan.FromSeconds(Math.Min(8, 2 * retryIndex));
}
