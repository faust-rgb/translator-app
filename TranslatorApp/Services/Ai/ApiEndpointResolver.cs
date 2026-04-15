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

    public static Uri ResolveOpenAiChatCompletionsUri(string baseUrl) =>
        ResolveOpenAiChatCompletionsUris(baseUrl).First();

    public static IReadOnlyList<Uri> ResolveOpenAiChatCompletionsUris(string baseUrl)
    {
        var trimmed = baseUrl.Trim().TrimEnd('/');
        if (trimmed.EndsWith("/v1/chat/completions", StringComparison.OrdinalIgnoreCase) ||
            trimmed.EndsWith("/v2/chat/completions", StringComparison.OrdinalIgnoreCase) ||
            trimmed.EndsWith("/v3/chat/completions", StringComparison.OrdinalIgnoreCase) ||
            trimmed.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
        {
            return [new Uri(trimmed)];
        }

        if (trimmed.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) ||
            trimmed.EndsWith("/v2", StringComparison.OrdinalIgnoreCase) ||
            trimmed.EndsWith("/v3", StringComparison.OrdinalIgnoreCase))
        {
            return
            [
                new Uri($"{trimmed}/chat/completions")
            ];
        }

        var candidates = new List<string>
        {
            $"{trimmed}/chat/completions",
            $"{trimmed}/v1/chat/completions",
            $"{trimmed}/v2/chat/completions",
            $"{trimmed}/v3/chat/completions"
        };

        return candidates
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(x => new Uri(x))
            .ToList();
    }
}
