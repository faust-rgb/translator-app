using CommunityToolkit.Mvvm.ComponentModel;

namespace TranslatorApp.Models;

public partial class DocumentTranslationItem : ObservableObject
{
    public Guid Id { get; } = Guid.NewGuid();

    [ObservableProperty]
    private string sourcePath = string.Empty;

    [ObservableProperty]
    private string outputPath = string.Empty;

    [ObservableProperty]
    private string fileType = string.Empty;

    [ObservableProperty]
    private int progress;

    [ObservableProperty]
    private string progressText = "待处理";

    [ObservableProperty]
    private DocumentStatus status = DocumentStatus.Pending;

    [ObservableProperty]
    private string? errorMessage;
}
