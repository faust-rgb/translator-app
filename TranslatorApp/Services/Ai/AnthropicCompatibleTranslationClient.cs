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
        using var message = new HttpRequestMessage(HttpMethod.Post, uri);
        message.Headers.Add("x-api-key", settings.ApiKey);
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
        message.Headers.Add("anthropic-version", settings.AnthropicVersion);
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        AddCustomHeaders(message, settings.CustomHeaders);

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

        message.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
        httpClient.Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds);

        if (request.EnableStreaming && request.OnPartialResponse is not null)
        {
            return await ReadStreamAsync(message, request, cancellationToken);
        }

        using var response = await httpClient.SendAsync(message, cancellationToken);
        await EnsureSuccessWithDetailsAsync(response, cancellationToken);
        var json = await response.Content.ReadAsStringAsync(cancellationToken);
        using var document = JsonDocument.Parse(json);
        return document.RootElement.GetProperty("content")[0].GetProperty("text").GetString()?.Trim() ?? string.Empty;
    }

    private async Task<string> ReadStreamAsync(HttpRequestMessage message, TranslationRequest request, CancellationToken cancellationToken)
    {
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
}
