using System.Net.Http;
using Microsoft.Extensions.DependencyInjection;
using TranslatorApp.Configuration;

namespace TranslatorApp.Services.Ai;

public sealed class AiTranslationClientFactory(
    IHttpClientFactory httpClientFactory,
    ITranslationPromptBuilder promptBuilder) : IAiTranslationClientFactory
{
    public IAiTranslationClient Create(AiSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        return settings.ProviderType switch
        {
            "OpenAiCompatible" => new OpenAiCompatibleTranslationClient(httpClientFactory.CreateClient(nameof(OpenAiCompatibleTranslationClient)), promptBuilder, settings),
            _ => new AnthropicCompatibleTranslationClient(httpClientFactory.CreateClient(nameof(AnthropicCompatibleTranslationClient)), promptBuilder, settings)
        };
    }
}
