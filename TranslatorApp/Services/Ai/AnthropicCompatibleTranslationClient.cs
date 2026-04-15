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
                using var document = JsonDocument.Parse(json);
                return document.RootElement.GetProperty("content")[0].GetProperty("text").GetString()?.Trim() ?? string.Empty;
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

            using var document = JsonDocument.Parse(payload);
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
