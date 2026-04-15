namespace TranslatorApp.Configuration;

public sealed class AiSettings
{
    public string ProviderType { get; set; } = "AnthropicCompatible";
    public string BaseUrl { get; set; } = "https://api.anthropic.com";
    public string ApiKey { get; set; } = string.Empty;
    public string Model { get; set; } = "claude-3-7-sonnet-latest";
    public string AnthropicVersion { get; set; } = "2023-06-01";
    public double Temperature { get; set; } = 0.1;
    public int MaxTokens { get; set; } = 4096;
    public string CustomHeaders { get; set; } = string.Empty;
    public int TimeoutSeconds { get; set; } = 300;
}
