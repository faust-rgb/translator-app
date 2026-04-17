using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Diagnostics;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;
using TranslatorApp.Services;
using TranslatorApp.Services.Documents;

namespace TranslatorApp.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private const int MaxLogLines = 400;
    private const int MaxStreamPreviewCharacters = 12000;
    private readonly ISettingsService _settingsService;
    private readonly IDocumentTranslationCoordinator _coordinator;
    private readonly IAppLogService _logService;
    private readonly ITranslationProgressService _translationProgressService;
    private readonly ITranslationHistoryService _historyService;
    private readonly IRecoveryStateService _recoveryStateService;
    private readonly IConnectionTestService _connectionTestService;
    private CancellationTokenSource? _translationCts;
    private CancellationTokenSource? _connectionTestCts;

    [ObservableProperty]
    private string providerType = "AnthropicCompatible";

    [ObservableProperty]
    private string baseUrl = string.Empty;

    [ObservableProperty]
    private string apiKey = string.Empty;

    [ObservableProperty]
    private string model = string.Empty;

    [ObservableProperty]
    private string anthropicVersion = "2023-06-01";

    [ObservableProperty]
    private string customHeaders = string.Empty;

    [ObservableProperty]
    private string sourceLanguage = "自动检测";

    [ObservableProperty]
    private string targetLanguage = "中文";

    [ObservableProperty]
    private string outputDirectory = string.Empty;

    [ObservableProperty]
    private string outputFontFamily = PdfSharpFontResolver.DefaultFontFamily;

    [ObservableProperty]
    private double outputFontSize = 11;

    [ObservableProperty]
    private int maxParallelDocuments = 1;

    [ObservableProperty]
    private int maxParallelBlocks = 1;

    [ObservableProperty]
    private int maxGlobalTranslationRequests = 1;

    [ObservableProperty]
    private double temperature = 0.1;

    [ObservableProperty]
    private int maxTokens = 4096;

    [ObservableProperty]
    private string glossaryPath = string.Empty;

    [ObservableProperty]
    private bool exportBilingualDocument = true;

    [ObservableProperty]
    private bool enableStreaming = true;

    [ObservableProperty]
    private int retryCount = 2;

    [ObservableProperty]
    private bool enableOcrForScannedPdf = true;

    [ObservableProperty]
    private string tesseractDataPath = string.Empty;

    [ObservableProperty]
    private string ocrLanguage = "chi_sim+eng";

    [ObservableProperty]
    private bool isRunning;

    [ObservableProperty]
    private bool isPaused;

    public string PauseButtonText => IsPaused ? "恢复任务" : "暂停任务";

    [ObservableProperty]
    private int overallProgress;

    [ObservableProperty]
    private string logText = string.Empty;

    [ObservableProperty]
    private string streamPreviewTitle = "实时译文预览";

    [ObservableProperty]
    private string streamPreviewText = string.Empty;

    [ObservableProperty]
    private bool isTestingConnection;

    [ObservableProperty]
    private DocumentTranslationItem? selectedDocument;

    public ObservableCollection<DocumentTranslationItem> Documents { get; } = [];
    public ObservableCollection<TranslationHistoryRecord> HistoryItems { get; } = [];
    public IReadOnlyList<string> CommonLanguages { get; } =
    [
        "自动检测",
        "中文",
        "英语",
        "日语",
        "韩语",
        "法语",
        "德语",
        "西班牙语",
        "葡萄牙语",
        "俄语",
        "阿拉伯语",
        "意大利语",
        "荷兰语",
        "泰语",
        "越南语",
        "印尼语",
        "土耳其语",
        "印地语"
    ];

    partial void OnTemperatureChanged(double value)
    {
        var clamped = Math.Clamp(value, 0, 2);
        if (clamped != value)
        {
            temperature = clamped;
        }
    }

    partial void OnMaxTokensChanged(int value)
    {
        var clamped = Math.Max(1, value);
        if (clamped != value)
        {
            maxTokens = clamped;
        }
    }

    partial void OnOutputFontSizeChanged(double value)
    {
        var clamped = Math.Clamp(value, 6, 72);
        if (clamped != value)
        {
            outputFontSize = clamped;
        }
    }

    partial void OnMaxParallelDocumentsChanged(int value)
    {
        var clamped = Math.Max(1, value);
        if (clamped != value)
        {
            maxParallelDocuments = clamped;
        }
    }

    partial void OnMaxParallelBlocksChanged(int value)
    {
        var clamped = Math.Max(1, value);
        if (clamped != value)
        {
            maxParallelBlocks = clamped;
        }
    }

    partial void OnMaxGlobalTranslationRequestsChanged(int value)
    {
        var clamped = Math.Max(1, value);
        if (clamped != value)
        {
            maxGlobalTranslationRequests = clamped;
        }
    }

    partial void OnRetryCountChanged(int value)
    {
        var clamped = Math.Clamp(value, 0, 10);
        if (clamped != value)
        {
            retryCount = clamped;
        }
    }

    public MainViewModel(
        ISettingsService settingsService,
        IDocumentTranslationCoordinator coordinator,
        IAppLogService logService,
        ITranslationProgressService translationProgressService,
        ITranslationHistoryService historyService,
        IRecoveryStateService recoveryStateService,
        IConnectionTestService connectionTestService)
    {
        _settingsService = settingsService;
        _coordinator = coordinator;
        _logService = logService;
        _translationProgressService = translationProgressService;
        _historyService = historyService;
        _recoveryStateService = recoveryStateService;
        _connectionTestService = connectionTestService;
        _logService.LogAdded += (_, line) =>
        {
            LogText = AppendLogLine(LogText, line);
        };
        _translationProgressService.StreamUpdated += (_, args) =>
        {
            StreamPreviewTitle = string.IsNullOrWhiteSpace(args.Title) ? "实时译文预览" : $"实时译文预览 - {args.Title}";
            StreamPreviewText = TrimToMaxCharacters(args.PartialText, MaxStreamPreviewCharacters);
        };
        Documents.CollectionChanged += OnDocumentsCollectionChanged;
    }

    public async Task InitializeAsync()
    {
        var settings = await _settingsService.LoadAsync();
        ProviderType = settings.Ai.ProviderType;
        BaseUrl = settings.Ai.BaseUrl;
        ApiKey = settings.Ai.ApiKey;
        Model = settings.Ai.Model;
        AnthropicVersion = settings.Ai.AnthropicVersion;
        CustomHeaders = settings.Ai.CustomHeaders;
        SourceLanguage = settings.Translation.SourceLanguage;
        TargetLanguage = settings.Translation.TargetLanguage;
        OutputDirectory = settings.Translation.OutputDirectory;
        OutputFontFamily = settings.Translation.OutputFontFamily;
        OutputFontSize = settings.Translation.OutputFontSize;
        MaxParallelDocuments = settings.Translation.MaxParallelDocuments;
        MaxParallelBlocks = settings.Translation.MaxParallelBlocks;
        MaxGlobalTranslationRequests = settings.Translation.MaxGlobalTranslationRequests;
        GlossaryPath = settings.Translation.GlossaryPath;
        ExportBilingualDocument = settings.Translation.ExportBilingualDocument;
        EnableStreaming = settings.Translation.EnableStreaming;
        RetryCount = settings.Translation.RetryCount;
        EnableOcrForScannedPdf = settings.Ocr.EnableOcrForScannedPdf;
        TesseractDataPath = settings.Ocr.TesseractDataPath;
        OcrLanguage = settings.Ocr.Language;
        Temperature = settings.Ai.Temperature;
        MaxTokens = settings.Ai.MaxTokens;

        var history = await _historyService.LoadAsync();
        HistoryItems.Clear();
        foreach (var item in history)
        {
            HistoryItems.Add(item);
        }

        var pending = await _recoveryStateService.LoadPendingAsync();
        foreach (var checkpoint in pending)
        {
            if (Documents.Any(x => string.Equals(x.SourcePath, checkpoint.SourcePath, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            Documents.Add(new DocumentTranslationItem
            {
                SourcePath = checkpoint.SourcePath,
                OutputPath = checkpoint.OutputPath,
                FileType = checkpoint.FileType,
                Progress = checkpoint.Progress,
                ProgressText = $"已恢复：{checkpoint.ProgressText}",
                Status = DocumentStatus.Pending,
                ErrorMessage = checkpoint.ErrorMessage
            });
        }

        if (pending.Count > 0)
        {
            _logService.Info($"已恢复 {pending.Count} 个未完成任务，可直接点击“开始翻译”继续。");
        }
    }

    [RelayCommand]
    private async Task LoadedAsync() => await InitializeAsync();

    [RelayCommand]
    private void AddFiles()
    {
        var dialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "支持的文档|*.docx;*.xlsx;*.pptx;*.pdf|全部文件|*.*"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (var file in dialog.FileNames.Where(file => Documents.All(x => !string.Equals(x.SourcePath, file, StringComparison.OrdinalIgnoreCase))))
        {
            Documents.Add(new DocumentTranslationItem
            {
                SourcePath = file,
                FileType = Path.GetExtension(file).TrimStart('.').ToUpperInvariant()
            });
        }

        _logService.Info($"已加入 {dialog.FileNames.Length} 个文件。");
    }

    [RelayCommand]
    private void ClearFiles() => Documents.Clear();

    [RelayCommand]
    private async Task RemoveSelectedAsync(DocumentTranslationItem? item)
    {
        if (item is not null)
        {
            Documents.Remove(item);
            await _recoveryStateService.RemoveCheckpointAsync(item.SourcePath);
            _logService.Info($"已删除任务：{Path.GetFileName(item.SourcePath)}");
        }
    }

    [RelayCommand]
    private void OpenOutputDocument(DocumentTranslationItem? item)
    {
        var target = item ?? SelectedDocument;
        if (target is null)
        {
            _logService.Error("未选择任务，无法打开译后文档。");
            return;
        }

        if (string.IsNullOrWhiteSpace(target.OutputPath) || !File.Exists(target.OutputPath))
        {
            _logService.Error("译后文档不存在，无法打开。");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = target.OutputPath,
            UseShellExecute = true
        });
        _logService.Info($"已打开译后文档：{Path.GetFileName(target.OutputPath)}");
    }

    [RelayCommand]
    private async Task ResumeTaskAsync(DocumentTranslationItem? item)
    {
        if (item is null)
        {
            return;
        }

        var checkpoint = await _recoveryStateService.GetCheckpointAsync(item.SourcePath);
        if (checkpoint is null)
        {
            _logService.Info($"任务 {Path.GetFileName(item.SourcePath)} 没有找到检查点，将从头开始。");
        }
        else
        {
            var progressInfo = checkpoint.UnitIndex > 0
                ? $"已完成约 {checkpoint.Progress}%（单元 {checkpoint.UnitIndex}"
                : $"已完成约 {checkpoint.Progress}%";
            var subInfo = checkpoint.SubUnitIndex > 0
                ? $", 子单元 {checkpoint.SubUnitIndex}）"
                : "）";
            _logService.Info($"任务 {Path.GetFileName(item.SourcePath)} 将从检查点继续：{progressInfo}{subInfo}，上次更新：{checkpoint.UpdatedAt:yyyy-MM-dd HH:mm:ss}");
        }

        item.Status = DocumentStatus.Pending;
        item.ErrorMessage = null;
        item.ProgressText = item.Progress > 0 ? $"等待继续：{checkpoint?.ProgressText ?? item.ProgressText}" : "等待开始";
        await _recoveryStateService.SaveCheckpointAsync(new DocumentCheckpoint
        {
            SourcePath = item.SourcePath,
            OutputPath = item.OutputPath,
            FileType = item.FileType,
            Status = nameof(DocumentStatus.Pending),
            Progress = item.Progress,
            ProgressText = item.ProgressText,
            UnitIndex = checkpoint?.UnitIndex ?? 0,
            SubUnitIndex = checkpoint?.SubUnitIndex ?? 0,
            ErrorMessage = null,
            UpdatedAt = DateTime.UtcNow
        });
        _logService.Info($"任务已设为继续：{Path.GetFileName(item.SourcePath)}");
    }

    [RelayCommand]
    private void PickOutputDirectory()
    {
        var dialog = new VistaFolderBrowserDialog();
        if (dialog.ShowDialog() == true)
        {
            OutputDirectory = dialog.SelectedPath;
        }
    }

    [RelayCommand]
    private void PickGlossaryFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "术语表|*.txt;*.tsv;*.csv|全部文件|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            GlossaryPath = dialog.FileName;
        }
    }

    [RelayCommand]
    private void PickTesseractDataDirectory()
    {
        var dialog = new VistaFolderBrowserDialog();
        if (dialog.ShowDialog() == true)
        {
            TesseractDataPath = dialog.SelectedPath;
        }
    }

    [RelayCommand]
    private async Task LoadConfigFileAsync()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON 配置文件|*.json|全部文件|*.*"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var settings = await _settingsService.LoadFromFileAsync(dialog.FileName);
            ProviderType = string.IsNullOrWhiteSpace(settings.Ai.ProviderType) ? ProviderType : settings.Ai.ProviderType;
            Model = settings.Ai.Model;
            BaseUrl = settings.Ai.BaseUrl;
            ApiKey = settings.Ai.ApiKey;
            _logService.Info($"已加载配置文件：{Path.GetFileName(dialog.FileName)}");
        }
        catch (Exception ex)
        {
            _logService.Error($"加载配置文件失败：{ex.Message}");
        }
    }

    [RelayCommand]
    private async Task ExportConfigFileAsync()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "JSON 配置文件|*.json|全部文件|*.*",
            FileName = "ai-config.json",
            DefaultExt = ".json"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var settings = new AppSettings
            {
                Ai = new AiSettings
                {
                    ProviderType = ProviderType,
                    BaseUrl = BaseUrl,
                    ApiKey = ApiKey,
                    Model = Model
                }
            };

            await _settingsService.SaveToFileAsync(settings, dialog.FileName);
            _logService.Info($"已导出配置文件：{Path.GetFileName(dialog.FileName)}");
        }
        catch (Exception ex)
        {
            _logService.Error($"导出配置文件失败：{ex.Message}");
        }
    }

    [RelayCommand]
    private async Task SaveSettingsAsync()
    {
        var settings = BuildSettings();
        await _settingsService.SaveAsync(settings);
        _logService.Info("配置已保存。");
    }

    [RelayCommand]
    private async Task ClearHistoryAsync()
    {
        await _historyService.ClearAsync();
        HistoryItems.Clear();
        _logService.Info("任务历史已清空。");
    }

    [RelayCommand]
    private async Task TestConnectionAsync()
    {
        if (IsTestingConnection)
        {
            return;
        }

        var settings = BuildSettings();
        await _settingsService.SaveAsync(settings);

        IsTestingConnection = true;
        _connectionTestCts?.Cancel();
        _connectionTestCts = new CancellationTokenSource(TimeSpan.FromSeconds(60));

        try
        {
            _logService.Info($"开始测试连接：{ProviderType} {BaseUrl} 模型 {Model}");
            await _connectionTestService.TestAsync(settings, _connectionTestCts.Token);
        }
        catch (OperationCanceledException)
        {
            _logService.Error("连接测试已取消或超时。");
        }
        catch (Exception ex)
        {
            _logService.Error($"连接测试失败：{ex.Message}");
        }
        finally
        {
            IsTestingConnection = false;
        }
    }

    [RelayCommand]
    private async Task StartAsync()
    {
        if (IsRunning || Documents.Count == 0)
        {
            return;
        }

        var settings = BuildSettings();
        await _settingsService.SaveAsync(settings);

        IsRunning = true;
        IsPaused = false;
        OverallProgress = 0;
        StreamPreviewText = string.Empty;
        StreamPreviewTitle = "实时译文预览";
        _translationProgressService.Clear();
        _coordinator.PauseController.Resume();
        _translationCts = new CancellationTokenSource();

        try
        {
            _logService.Info("开始执行批量翻译任务。");
            await _coordinator.RunAsync(Documents, settings, _translationCts.Token);
        }
        finally
        {
            IsRunning = false;
            IsPaused = false;
            OverallProgress = Documents.Count == 0 ? 0 : (int)Math.Round(Documents.Average(x => x.Progress));
            _ = RefreshHistoryAsync();
        }
    }

    [RelayCommand]
    private void PauseOrResume()
    {
        if (!IsRunning)
        {
            return;
        }

        if (IsPaused)
        {
            _coordinator.PauseController.Resume();
            IsPaused = false;
            _logService.Info("任务已恢复。");
        }
        else
        {
            _coordinator.PauseController.Pause();
            IsPaused = true;
            _logService.Info("任务已暂停。");
        }
    }

    [RelayCommand]
    private void Stop()
    {
        _translationCts?.Cancel();
        _coordinator.PauseController.Resume();
        IsPaused = false;
        _logService.Info("已请求停止当前任务。");
    }

    partial void OnLogTextChanged(string value)
    {
        RefreshOverallProgress();
    }

    partial void OnIsPausedChanged(bool value) => OnPropertyChanged(nameof(PauseButtonText));

    private void OnDocumentsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (DocumentTranslationItem item in e.NewItems)
            {
                item.PropertyChanged += OnDocumentItemPropertyChanged;
            }
        }

        if (e.OldItems is not null)
        {
            foreach (DocumentTranslationItem item in e.OldItems)
            {
                item.PropertyChanged -= OnDocumentItemPropertyChanged;
            }
        }

        RefreshOverallProgress();
    }

    private void OnDocumentItemPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(DocumentTranslationItem.Progress) or nameof(DocumentTranslationItem.Status))
        {
            RefreshOverallProgress();
        }
    }

    private void RefreshOverallProgress()
    {
        OverallProgress = Documents.Count == 0 ? 0 : (int)Math.Round(Documents.Average(x => x.Progress));
    }

    private AppSettings BuildSettings() =>
        new()
        {
            Ai = new AiSettings
            {
                ProviderType = ProviderType,
                BaseUrl = BaseUrl,
                ApiKey = ApiKey,
                Model = Model,
                AnthropicVersion = AnthropicVersion,
                CustomHeaders = CustomHeaders,
                Temperature = Math.Clamp(Temperature, 0, 2),
                MaxTokens = Math.Max(1, MaxTokens)
            },
            Translation = new TranslationSettings
            {
                SourceLanguage = SourceLanguage,
                TargetLanguage = TargetLanguage,
                OutputDirectory = OutputDirectory,
                OutputFontFamily = OutputFontFamily,
                OutputFontSize = Math.Clamp(OutputFontSize, 6, 72),
                MaxParallelDocuments = Math.Max(1, MaxParallelDocuments),
                MaxParallelBlocks = Math.Max(1, MaxParallelBlocks),
                MaxGlobalTranslationRequests = Math.Max(1, MaxGlobalTranslationRequests),
                GlossaryPath = GlossaryPath,
                ExportBilingualDocument = ExportBilingualDocument,
                EnableStreaming = EnableStreaming,
                RetryCount = Math.Clamp(RetryCount, 0, 10)
            },
            Ocr = new Configuration.OcrSettings
            {
                EnableOcrForScannedPdf = EnableOcrForScannedPdf,
                TesseractDataPath = TesseractDataPath,
                Language = OcrLanguage
            }
        };

    private async Task RefreshHistoryAsync()
    {
        var history = await _historyService.LoadAsync();
        HistoryItems.Clear();
        foreach (var item in history)
        {
            HistoryItems.Add(item);
        }
    }

    private static string AppendLogLine(string current, string nextLine)
    {
        if (string.IsNullOrWhiteSpace(nextLine))
        {
            return current;
        }

        var combined = string.IsNullOrWhiteSpace(current)
            ? nextLine
            : string.Concat(current, Environment.NewLine, nextLine);

        var lineCount = combined.AsSpan().Count('\n');

        if (lineCount < MaxLogLines)
        {
            return combined;
        }

        var linesToSkip = lineCount - MaxLogLines + 1;
        var skipped = 0;
        var cutIndex = 0;
        for (var i = 0; i < combined.Length; i++)
        {
            if (combined[i] == '\n')
            {
                skipped++;
                if (skipped >= linesToSkip)
                {
                    cutIndex = i + 1;
                    break;
                }
            }
        }

        return cutIndex < combined.Length ? combined[cutIndex..] : combined;
    }

    private static string TrimToMaxCharacters(string? value, int maxCharacters)
    {
        var text = value ?? string.Empty;
        if (text.Length <= maxCharacters)
        {
            return text;
        }

        return text[^maxCharacters..];
    }
}
