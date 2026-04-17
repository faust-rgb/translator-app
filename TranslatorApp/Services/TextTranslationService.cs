using TranslatorApp.Configuration;
using TranslatorApp.Models;
using TranslatorApp.Services.Ai;

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

        var client = clientFactory.Create(settings.Ai);
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

        var retryCount = Math.Max(0, settings.Translation.RetryCount);
        var exceptions = new List<Exception>();
        for (var attempt = 0; attempt <= retryCount; attempt++)
        {
            try
            {
                await using var _ = await translationRequestThrottle.AcquireAsync(
                    settings.Translation.MaxGlobalTranslationRequests,
                    cancellationToken);
                return await client.TranslateAsync(request, cancellationToken);
            }
            catch (Exception ex) when (attempt < retryCount && ex is not OperationCanceledException)
            {
                exceptions.Add(ex);
                logService.Error($"翻译请求失败，第 {attempt + 1} 次重试前等待：{ex.Message}");
                await Task.Delay(TimeSpan.FromSeconds(Math.Min(8, 2 * (attempt + 1))), cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                exceptions.Add(ex);
                break;
            }
        }

        if (exceptions.Count == 0)
        {
            throw new InvalidOperationException("翻译失败。");
        }

        if (exceptions.Count == 1)
        {
            throw exceptions[0];
        }

        throw new AggregateException($"翻译在 {exceptions.Count} 次尝试后仍失败。", exceptions);
    }
}
