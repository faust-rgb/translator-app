namespace TranslatorApp.Models;

public sealed class TranslationRequest
{
    public required string Text { get; init; }
    public required string SourceLanguage { get; init; }
    public required string TargetLanguage { get; init; }
    public required string ContextHint { get; init; }
    public required string PromptTemplate { get; init; }
    public string GlossaryPrompt { get; init; } = string.Empty;
    public bool EnableStreaming { get; init; }
    public Func<string, Task>? OnPartialResponse { get; init; }
}
