namespace TranslatorApp.Configuration;

public sealed class AppSettings
{
    public AiSettings Ai { get; set; } = new();
    public TranslationSettings Translation { get; set; } = new();
    public OcrSettings Ocr { get; set; } = new();
}
