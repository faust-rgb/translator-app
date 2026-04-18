namespace TranslatorApp.Configuration;

public sealed class OcrSettings
{
    public bool EnableOcrForScannedPdf { get; set; } = true;
    public string TesseractDataPath { get; set; } = string.Empty;
    public string Language { get; set; } = "chi_sim+eng";
    public double RenderScale { get; set; } = 2.0;
    public int MinimumNativeTextWords { get; set; } = 8;
    public double SparseTextCoverageThreshold { get; set; } = 0.025;
    public int SparseTextBlockThreshold { get; set; } = 3;
    public float MinimumAcceptedConfidence { get; set; } = 35;
    public double OcrBlockMergeVerticalGapRatio { get; set; } = 0.95;
    public double OcrBlockMergeIndentationRatio { get; set; } = 1.2;
    public double OcrBlockMergeWidthDifferenceRatio { get; set; } = 0.55;
    public double OcrColumnCenterToleranceRatio { get; set; } = 0.12;
    public double OcrColumnLeftToleranceRatio { get; set; } = 0.08;
    public double OcrColumnOverlapThreshold { get; set; } = 0.35;
}
