using TranslatorApp.Configuration;
using TranslatorApp.Models;
using TranslatorApp.Services.Ai;

namespace TranslatorApp.Services;

public sealed class ConnectionTestService(
    IAiTranslationClientFactory clientFactory,
    IAppLogService logService) : IConnectionTestService
{
    public async Task TestAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(settings);

        var client = clientFactory.Create(settings.Ai);
        if (settings.Ai.ProviderType == "OpenAiCompatible")
        {
            var endpoints = Ai.ApiEndpointResolver.ResolveOpenAiChatCompletionsUris(settings.Ai.BaseUrl);
            logService.Info($"连接测试候选地址：{string.Join(" | ", endpoints)}");
        }
        else
        {
            var endpoint = Ai.ApiEndpointResolver.ResolveAnthropicMessagesUri(settings.Ai.BaseUrl);
            logService.Info($"连接测试目标地址：{endpoint}");
        }

        logService.Info($"连接测试模型：{settings.Ai.Model}");

        var request = new TranslationRequest
        {
            Text = "Hello",
            SourceLanguage = "英语",
            TargetLanguage = "中文",
            ContextHint = "连接测试",
            PromptTemplate = "请将下面文本翻译成中文，只返回译文。",
            GlossaryPrompt = string.Empty,
            EnableStreaming = false,
            OnPartialResponse = null
        };

        var result = await client.TranslateAsync(request, cancellationToken);
        if (string.IsNullOrWhiteSpace(result))
        {
            throw new InvalidOperationException("接口返回成功，但内容为空。");
        }

        logService.Info($"连接测试成功，返回示例：{result}");
    }
}
