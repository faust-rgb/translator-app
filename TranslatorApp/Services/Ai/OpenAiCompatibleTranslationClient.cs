using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using TranslatorApp.Configuration;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Ai;

public sealed class OpenAiCompatibleTranslationClient(
    HttpClient httpClient,
    ITranslationPromptBuilder promptBuilder,
    AiSettings settings) : IAiTranslationClient
{
    public async Task<string> TranslateAsync(TranslationRequest request, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = settings.Model,
            temperature = settings.Temperature,
            stream = request.EnableStreaming,
            messages = new object[]
            {
                new
                {
                    role = "system",
                    content = promptBuilder.BuildSystemPrompt(request)
                },
                new
                {
                    role = "user",
                    content = promptBuilder.BuildUserPrompt(request)
                }
            }
        };

        httpClient.Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds);
        var content = JsonSerializer.Serialize(payload);
        Exception? lastException = null;
        var attemptedEndpoints = new List<string>();

        foreach (var uri in ApiEndpointResolver.ResolveOpenAiChatCompletionsUris(settings.BaseUrl))
        {
            try
            {
                if (request.EnableStreaming && request.OnPartialResponse is not null)
                {
                    return await ReadStreamAsync(CreateMessage(uri, content), request, cancellationToken);
                }

                using var message = CreateMessage(uri, content);
                using var response = await httpClient.SendAsync(message, cancellationToken);
                await EnsureSuccessWithDetailsAsync(response, cancellationToken);
                var json = await response.Content.ReadAsStringAsync(cancellationToken);
                using var document = ParseJsonWithDetails(json);
                return document.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString()?.Trim() ?? string.Empty;
            }
            catch (Exception ex) when (IsEndpointCompatibilityFailure(ex))
            {
                lastException = ex;
                attemptedEndpoints.Add($"{uri} -> {ex.Message}");
            }
        }

        if (lastException is null)
        {
            throw new InvalidOperationException("OpenAI 兼容请求失败。");
        }

        var attempts = attemptedEndpoints.Count == 0
            ? "未记录到候选端点。"
            : "已尝试端点：" + Environment.NewLine + string.Join(Environment.NewLine, attemptedEndpoints.Select(x => $"- {x}"));
        throw new InvalidOperationException($"OpenAI 兼容请求失败。{Environment.NewLine}{attempts}", lastException);
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

            using var document = ParseJsonWithDetails(payload);
            if (document.RootElement.TryGetProperty("choices", out var choicesElement) &&
                choicesElement.GetArrayLength() > 0 &&
                choicesElement[0].TryGetProperty("delta", out var deltaElement) &&
                deltaElement.TryGetProperty("content", out var contentElement))
            {
                var chunk = contentElement.GetString();
                if (!string.IsNullOrEmpty(chunk))
                {
                    builder.Append(chunk);
                    await request.OnPartialResponse!(builder.ToString());
                }
            }
        }

        return builder.ToString().Trim();
    }

    private HttpRequestMessage CreateMessage(Uri uri, string content)
    {
        var message = new HttpRequestMessage(HttpMethod.Post, uri);
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        message.Content = new StringContent(content, Encoding.UTF8, "application/json");
        return message;
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

    private static JsonDocument ParseJsonWithDetails(string content)
    {
        try
        {
            return JsonDocument.Parse(content);
        }
        catch (JsonException ex)
        {
            var snippet = string.IsNullOrWhiteSpace(content) ? "空响应" : content.Trim();
            if (snippet.Length > 600)
            {
                snippet = snippet[..600];
            }

            throw new InvalidOperationException($"响应不是有效 JSON。原始内容片段：{snippet}", ex);
        }
    }

    private static bool IsEndpointCompatibilityFailure(Exception ex) =>
        ex.Message.Contains("404", StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("405", StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("响应不是有效 JSON", StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("<html", StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("<!DOCTYPE", StringComparison.OrdinalIgnoreCase);
}
