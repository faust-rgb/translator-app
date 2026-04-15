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
        var uri = ApiEndpointResolver.ResolveOpenAiChatCompletionsUri(settings.BaseUrl);
        using var message = new HttpRequestMessage(HttpMethod.Post, uri);
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

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
        return document.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString()?.Trim() ?? string.Empty;
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
}
