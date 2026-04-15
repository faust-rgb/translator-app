namespace TranslatorApp.Models;

public sealed class TranslationHistoryRecord
{
    public DateTime Timestamp { get; set; }
    public string SourcePath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public string ProviderType { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public int Progress { get; set; }
    public double DurationSeconds { get; set; }
    public string? ErrorMessage { get; set; }
}
