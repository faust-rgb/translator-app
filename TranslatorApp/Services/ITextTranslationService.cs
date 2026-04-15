using TranslatorApp.Configuration;

namespace TranslatorApp.Services;

public interface ITextTranslationService
{
    Task<string> TranslateAsync(
        string text,
        string contextHint,
        AppSettings settings,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken);
}
