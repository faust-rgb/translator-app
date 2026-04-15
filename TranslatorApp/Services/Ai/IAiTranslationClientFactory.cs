using TranslatorApp.Configuration;

namespace TranslatorApp.Services.Ai;

public interface IAiTranslationClientFactory
{
    IAiTranslationClient Create(AiSettings settings);
}
