namespace TranslatorApp.Configuration;

public sealed class OcrSettings
{
    public bool EnableOcrForScannedPdf { get; set; } = true;
    public string TesseractDataPath { get; set; } = string.Empty;
    public string Language { get; set; } = "chi_sim+eng";
    public double RenderScale { get; set; } = 2.0;
    public int MinimumNativeTextWords { get; set; } = 8;
}
