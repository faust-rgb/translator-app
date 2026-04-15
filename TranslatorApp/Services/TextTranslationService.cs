using TranslatorApp.Configuration;
using TranslatorApp.Models;
using TranslatorApp.Services.Ai;

namespace TranslatorApp.Services;

public sealed class TextTranslationService(
    IAiTranslationClientFactory clientFactory,
    IGlossaryService glossaryService,
    IAppLogService logService) : ITextTranslationService
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
            EnableStreaming = settings.Translation.EnableStreaming,
            OnPartialResponse = onPartialResponse
        };

        var retryCount = Math.Max(0, settings.Translation.RetryCount);
        Exception? lastException = null;
        for (var attempt = 0; attempt <= retryCount; attempt++)
        {
            try
            {
                return await client.TranslateAsync(request, cancellationToken);
            }
            catch (Exception ex) when (attempt < retryCount && ex is not OperationCanceledException)
            {
                lastException = ex;
                logService.Error($"翻译请求失败，第 {attempt + 1} 次重试前等待：{ex.Message}");
                await Task.Delay(TimeSpan.FromSeconds(Math.Min(8, 2 * (attempt + 1))), cancellationToken);
            }
        }

        throw lastException ?? new InvalidOperationException("翻译失败。");
    }
}
