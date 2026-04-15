using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using TranslatorApp.Configuration;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Ai;

public sealed class AnthropicCompatibleTranslationClient(
    HttpClient httpClient,
    ITranslationPromptBuilder promptBuilder,
    AiSettings settings) : IAiTranslationClient
{
    public async Task<string> TranslateAsync(TranslationRequest request, CancellationToken cancellationToken)
    {
        var uri = ApiEndpointResolver.ResolveAnthropicMessagesUri(settings.BaseUrl);
        var payload = new
        {
            model = settings.Model,
            max_tokens = settings.MaxTokens,
            temperature = settings.Temperature,
            stream = request.EnableStreaming,
            system = promptBuilder.BuildSystemPrompt(request),
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = new[]
                    {
                        new
                        {
                            type = "text",
                            text = promptBuilder.BuildUserPrompt(request)
                        }
                    }
                }
            }
        };
        httpClient.Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds);

        var content = JsonSerializer.Serialize(payload);
        var authModes = new[]
        {
            AnthropicAuthMode.XApiKeyOnly,
            AnthropicAuthMode.BearerOnly,
            AnthropicAuthMode.XApiKeyAndBearer
        };

        Exception? lastException = null;
        foreach (var authMode in authModes)
        {
            try
            {
                if (request.EnableStreaming && request.OnPartialResponse is not null)
                {
                    return await ReadStreamAsync(uri, content, authMode, request, cancellationToken);
                }

                using var message = CreateMessage(uri, content, authMode);
                using var response = await httpClient.SendAsync(message, cancellationToken);
                await EnsureSuccessWithDetailsAsync(response, cancellationToken);
                var json = await response.Content.ReadAsStringAsync(cancellationToken);
                return ExtractTextFromResponse(json);
            }
            catch (HttpRequestException ex) when (IsAuthenticationFailure(ex))
            {
                lastException = ex;
            }
        }

        throw lastException ?? new InvalidOperationException("Anthropic 兼容请求失败。");
    }

    private async Task<string> ReadStreamAsync(
        Uri uri,
        string content,
        AnthropicAuthMode authMode,
        TranslationRequest request,
        CancellationToken cancellationToken)
    {
        using var message = CreateMessage(uri, content, authMode);
        using var response = await httpClient.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        await EnsureSuccessWithDetailsAsync(response, cancellationToken);
        using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var reader = new StreamReader(stream);

        var builder = new StringBuilder();
        while (true)
        {
            var line = await reader.ReadLineAsync(cancellationToken);
            if (line is null)
            {
                break;
            }

            if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("data:", StringComparison.Ordinal))
            {
                continue;
            }

            var payload = line["data:".Length..].Trim();
            if (payload == "[DONE]")
            {
                break;
            }

            using var document = ParseJsonWithDetails(payload);
            if (document.RootElement.TryGetProperty("type", out var typeElement) &&
                typeElement.GetString() == "content_block_delta" &&
                document.RootElement.TryGetProperty("delta", out var deltaElement) &&
                deltaElement.TryGetProperty("text", out var textElement))
            {
                var chunk = textElement.GetString();
                if (!string.IsNullOrEmpty(chunk))
                {
                    builder.Append(chunk);
                    await request.OnPartialResponse!(builder.ToString());
                }
                continue;
            }

            if (TryExtractLooseText(document.RootElement, out var looseChunk) && !string.IsNullOrWhiteSpace(looseChunk))
            {
                builder.Append(looseChunk);
                await request.OnPartialResponse!(builder.ToString());
            }
        }

        return builder.ToString().Trim();
    }

    private HttpRequestMessage CreateMessage(Uri uri, string content, AnthropicAuthMode authMode)
    {
        var message = new HttpRequestMessage(HttpMethod.Post, uri);
        ApplyAuthHeaders(message, authMode);
        message.Headers.Add("anthropic-version", settings.AnthropicVersion);
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        AddCustomHeaders(message, settings.CustomHeaders);
        message.Content = new StringContent(content, Encoding.UTF8, "application/json");
        return message;
    }

    private void ApplyAuthHeaders(HttpRequestMessage message, AnthropicAuthMode authMode)
    {
        switch (authMode)
        {
            case AnthropicAuthMode.XApiKeyOnly:
                message.Headers.Add("x-api-key", settings.ApiKey);
                break;
            case AnthropicAuthMode.BearerOnly:
                message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                break;
            case AnthropicAuthMode.XApiKeyAndBearer:
                message.Headers.Add("x-api-key", settings.ApiKey);
                message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                break;
        }
    }

    private static async Task EnsureSuccessWithDetailsAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        if (response.IsSuccessStatusCode)
        {
            return;
        }

        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        var snippet = string.IsNullOrWhiteSpace(body) ? "无响应正文" : body.Trim();
        if (snippet.Length > 600)
        {
            snippet = snippet[..600];
        }

        throw new HttpRequestException($"HTTP {(int)response.StatusCode} {response.ReasonPhrase}，响应内容：{snippet}");
    }

    private static string ExtractTextFromResponse(string json)
    {
        using var document = ParseJsonWithDetails(json);
        if (TryExtractLooseText(document.RootElement, out var text) && !string.IsNullOrWhiteSpace(text))
        {
            return text.Trim();
        }

        var snippet = json.Trim();
        if (snippet.Length > 800)
        {
            snippet = snippet[..800];
        }

        throw new InvalidOperationException($"响应 JSON 中未找到可识别的文本字段。原始内容片段：{snippet}");
    }

    private static bool TryExtractLooseText(JsonElement root, out string text)
    {
        if (root.ValueKind == JsonValueKind.Object)
        {
            if (root.TryGetProperty("content", out var contentElement))
            {
                if (contentElement.ValueKind == JsonValueKind.String)
                {
                    text = contentElement.GetString() ?? string.Empty;
                    return !string.IsNullOrWhiteSpace(text);
                }

                if (contentElement.ValueKind == JsonValueKind.Array)
                {
                    var parts = new List<string>();
                    foreach (var item in contentElement.EnumerateArray())
                    {
                        if (item.ValueKind == JsonValueKind.String)
                        {
                            var value = item.GetString();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                parts.Add(value);
                            }
                        }
                        else if (item.ValueKind == JsonValueKind.Object)
                        {
                            if (item.TryGetProperty("text", out var textElement))
                            {
                                var value = textElement.GetString();
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    parts.Add(value);
                                }
                            }
                            else if (item.TryGetProperty("content", out var nestedContent) &&
                                     nestedContent.ValueKind == JsonValueKind.String)
                            {
                                var value = nestedContent.GetString();
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    parts.Add(value);
                                }
                            }
                        }
                    }

                    text = string.Concat(parts);
                    return !string.IsNullOrWhiteSpace(text);
                }
            }

            if (root.TryGetProperty("completion", out var completionElement) &&
                completionElement.ValueKind == JsonValueKind.String)
            {
                text = completionElement.GetString() ?? string.Empty;
                return !string.IsNullOrWhiteSpace(text);
            }

            if (root.TryGetProperty("message", out var messageElement))
            {
                if (messageElement.ValueKind == JsonValueKind.String)
                {
                    text = messageElement.GetString() ?? string.Empty;
                    return !string.IsNullOrWhiteSpace(text);
                }

                if (messageElement.ValueKind == JsonValueKind.Object &&
                    messageElement.TryGetProperty("content", out var messageContent))
                {
                    if (messageContent.ValueKind == JsonValueKind.String)
                    {
                        text = messageContent.GetString() ?? string.Empty;
                        return !string.IsNullOrWhiteSpace(text);
                    }
                }
            }

            if (root.TryGetProperty("choices", out var choicesElement) &&
                choicesElement.ValueKind == JsonValueKind.Array &&
                choicesElement.GetArrayLength() > 0)
            {
                var choice = choicesElement[0];
                if (choice.ValueKind == JsonValueKind.Object)
                {
                    if (choice.TryGetProperty("message", out var choiceMessage) &&
                        choiceMessage.ValueKind == JsonValueKind.Object &&
                        choiceMessage.TryGetProperty("content", out var choiceContent))
                    {
                        if (choiceContent.ValueKind == JsonValueKind.String)
                        {
                            text = choiceContent.GetString() ?? string.Empty;
                            return !string.IsNullOrWhiteSpace(text);
                        }
                    }

                    if (choice.TryGetProperty("text", out var choiceText) &&
                        choiceText.ValueKind == JsonValueKind.String)
                    {
                        text = choiceText.GetString() ?? string.Empty;
                        return !string.IsNullOrWhiteSpace(text);
                    }
                }
            }
        }

        text = string.Empty;
        return false;
    }

    private static JsonDocument ParseJsonWithDetails(string content)
    {
        try
        {
            return JsonDocument.Parse(content);
        }
        catch (JsonException ex)
        {
            var snippet = string.IsNullOrWhiteSpace(content) ? "空响应" : content.Trim();
            if (snippet.Length > 800)
            {
                snippet = snippet[..800];
            }

            throw new InvalidOperationException($"响应不是有效 JSON。原始内容片段：{snippet}", ex);
        }
    }

    private static bool IsAuthenticationFailure(HttpRequestException exception) =>
        exception.Message.Contains("401", StringComparison.OrdinalIgnoreCase) ||
        exception.Message.Contains("403", StringComparison.OrdinalIgnoreCase) ||
        exception.Message.Contains("AuthenticationError", StringComparison.OrdinalIgnoreCase) ||
        exception.Message.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase);

    private static void AddCustomHeaders(HttpRequestMessage message, string rawHeaders)
    {
        if (string.IsNullOrWhiteSpace(rawHeaders))
        {
            return;
        }

        foreach (var line in rawHeaders.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var parts = line.Split(':', 2);
            if (parts.Length == 2)
            {
                message.Headers.TryAddWithoutValidation(parts[0].Trim(), parts[1].Trim());
            }
        }
    }

    private enum AnthropicAuthMode
    {
        XApiKeyOnly,
        BearerOnly,
        XApiKeyAndBearer
    }
}
