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
        var client = clientFactory.Create(settings.Ai);
        var endpoint = settings.Ai.ProviderType == "OpenAiCompatible"
            ? Ai.ApiEndpointResolver.ResolveOpenAiChatCompletionsUri(settings.Ai.BaseUrl)
            : Ai.ApiEndpointResolver.ResolveAnthropicMessagesUri(settings.Ai.BaseUrl);
        logService.Info($"连接测试目标地址：{endpoint}");

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
