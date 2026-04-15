namespace TranslatorApp.Services.Ai;

public static class ApiEndpointResolver
{
    public static Uri ResolveAnthropicMessagesUri(string baseUrl)
    {
        var trimmed = baseUrl.Trim().TrimEnd('/');
        if (trimmed.EndsWith("/v1/messages", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri(trimmed);
        }

        if (trimmed.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri($"{trimmed}/messages");
        }

        return new Uri($"{trimmed}/v1/messages");
    }

    public static Uri ResolveOpenAiChatCompletionsUri(string baseUrl)
    {
        var trimmed = baseUrl.Trim().TrimEnd('/');
        if (trimmed.EndsWith("/v1/chat/completions", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri(trimmed);
        }

        if (trimmed.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri($"{trimmed}/chat/completions");
        }

        return new Uri($"{trimmed}/v1/chat/completions");
    }
}
