namespace TranslatorApp.Models;

public sealed class DocumentCheckpoint
{
    public string SourcePath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public string FileType { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public int Progress { get; set; }
    public string ProgressText { get; set; } = string.Empty;
    public int UnitIndex { get; set; }
    public int SubUnitIndex { get; set; }
    public string? ErrorMessage { get; set; }
    public DateTime UpdatedAt { get; set; }
}
