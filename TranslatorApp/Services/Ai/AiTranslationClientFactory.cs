using System.Net.Http;
using Microsoft.Extensions.DependencyInjection;
using TranslatorApp.Configuration;

namespace TranslatorApp.Services.Ai;

public sealed class AiTranslationClientFactory(
    IHttpClientFactory httpClientFactory,
    ITranslationPromptBuilder promptBuilder) : IAiTranslationClientFactory
{
    public IAiTranslationClient Create(AiSettings settings) =>
        settings.ProviderType switch
        {
            "OpenAiCompatible" => new OpenAiCompatibleTranslationClient(httpClientFactory.CreateClient(), promptBuilder, settings),
            _ => new AnthropicCompatibleTranslationClient(httpClientFactory.CreateClient(), promptBuilder, settings)
        };
}
