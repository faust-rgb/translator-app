using TranslatorApp.Models;

namespace TranslatorApp.Services.Ai;

public interface IAiTranslationClient
{
    Task<string> TranslateAsync(TranslationRequest request, CancellationToken cancellationToken);
}
