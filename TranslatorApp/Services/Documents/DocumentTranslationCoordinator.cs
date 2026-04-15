using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public sealed class DocumentTranslationCoordinator(
    IEnumerable<IDocumentTranslator> translators,
    IAppLogService logService,
    ITranslationHistoryService historyService,
    IRecoveryStateService recoveryStateService) : IDocumentTranslationCoordinator
{
    public PauseController PauseController { get; } = new();

    public async Task RunAsync(ObservableCollection<DocumentTranslationItem> items, AppSettings settings, CancellationToken cancellationToken)
    {
        var maxParallel = Math.Max(1, settings.Translation.MaxParallelDocuments);
        using var semaphore = new SemaphoreSlim(maxParallel);

        var tasks = items
            .Where(x => x.Status is DocumentStatus.Pending)
            .Select(async item =>
            {
                await semaphore.WaitAsync(cancellationToken);
                try
                {
                    await TranslateSingleAsync(item, settings, cancellationToken);
                }
                finally
                {
                    semaphore.Release();
                }
            });

        await Task.WhenAll(tasks);
    }

    private async Task TranslateSingleAsync(DocumentTranslationItem item, AppSettings settings, CancellationToken cancellationToken)
    {
        var startedAt = DateTime.Now;
        var checkpoint = await recoveryStateService.GetCheckpointAsync(item.SourcePath);
        var extension = Path.GetExtension(item.SourcePath).ToLowerInvariant();
        var translator = translators.FirstOrDefault(x => x.CanHandle(extension));
        if (translator is null)
        {
            await UpdateItemAsync(item, () =>
            {
                item.Status = DocumentStatus.Failed;
                item.ErrorMessage = $"暂不支持 {extension} 文件。";
            });
            logService.Error($"{Path.GetFileName(item.SourcePath)} 不支持的格式：{extension}");
            await WriteHistoryAsync(item, settings, startedAt);
            await recoveryStateService.RemoveCheckpointAsync(item.SourcePath);
            return;
        }

        await UpdateItemAsync(item, () =>
        {
            item.Status = DocumentStatus.Running;
            item.Progress = 0;
            item.ProgressText = "准备开始";
            item.ErrorMessage = null;
        });
        await recoveryStateService.SaveCheckpointAsync(new DocumentCheckpoint
        {
            SourcePath = item.SourcePath,
            OutputPath = item.OutputPath,
            FileType = item.FileType,
            Status = nameof(DocumentStatus.Running),
            Progress = item.Progress,
            ProgressText = item.ProgressText,
            UnitIndex = checkpoint?.UnitIndex ?? 0,
            SubUnitIndex = checkpoint?.SubUnitIndex ?? 0,
            ErrorMessage = null,
            UpdatedAt = DateTime.Now
        });

        try
        {
            var context = new TranslationJobContext
            {
                Settings = settings,
                Item = item,
                CancellationToken = cancellationToken,
                PauseController = PauseController,
                ResumeUnitIndex = checkpoint?.UnitIndex ?? 0,
                ResumeSubUnitIndex = checkpoint?.SubUnitIndex ?? 0,
                ReportProgressAsync = (progress, text) => UpdateItemAsync(item, () =>
                {
                    item.Progress = progress;
                    item.ProgressText = text;
                    item.Status = PauseController.IsPaused ? DocumentStatus.Paused : DocumentStatus.Running;
                }),
                SaveCheckpointAsync = (unitIndex, subUnitIndex, text) => recoveryStateService.SaveCheckpointAsync(new DocumentCheckpoint
                {
                    SourcePath = item.SourcePath,
                    OutputPath = item.OutputPath,
                    FileType = item.FileType,
                    Status = PauseController.IsPaused ? nameof(DocumentStatus.Paused) : nameof(DocumentStatus.Running),
                    Progress = item.Progress,
                    ProgressText = text,
                    UnitIndex = unitIndex,
                    SubUnitIndex = subUnitIndex,
                    ErrorMessage = item.ErrorMessage,
                    UpdatedAt = DateTime.Now
                })
            };

            await translator.TranslateAsync(context);
            await UpdateItemAsync(item, () =>
            {
                item.Progress = 100;
                item.ProgressText = "已完成";
                item.Status = DocumentStatus.Completed;
            });
            logService.Info($"{Path.GetFileName(item.SourcePath)} 翻译完成。");
            await WriteHistoryAsync(item, settings, startedAt);
            await recoveryStateService.RemoveCheckpointAsync(item.SourcePath);
        }
        catch (OperationCanceledException)
        {
            await UpdateItemAsync(item, () =>
            {
                item.Status = DocumentStatus.Stopped;
                item.ProgressText = "已停止";
            });
            logService.Info($"{Path.GetFileName(item.SourcePath)} 已停止。");
            await WriteHistoryAsync(item, settings, startedAt);
            await recoveryStateService.SaveCheckpointAsync(new DocumentCheckpoint
            {
                SourcePath = item.SourcePath,
                OutputPath = item.OutputPath,
                FileType = item.FileType,
                Status = nameof(DocumentStatus.Stopped),
                Progress = item.Progress,
                ProgressText = item.ProgressText,
                UnitIndex = checkpoint?.UnitIndex ?? 0,
                SubUnitIndex = checkpoint?.SubUnitIndex ?? 0,
                ErrorMessage = item.ErrorMessage,
                UpdatedAt = DateTime.Now
            });
        }
        catch (Exception ex)
        {
            await UpdateItemAsync(item, () =>
            {
                item.Status = DocumentStatus.Failed;
                item.ErrorMessage = ex.Message;
                item.ProgressText = "失败";
            });
            logService.Error($"{Path.GetFileName(item.SourcePath)} 翻译失败：{ex.Message}");
            await WriteHistoryAsync(item, settings, startedAt);
            await recoveryStateService.SaveCheckpointAsync(new DocumentCheckpoint
            {
                SourcePath = item.SourcePath,
                OutputPath = item.OutputPath,
                FileType = item.FileType,
                Status = nameof(DocumentStatus.Failed),
                Progress = item.Progress,
                ProgressText = item.ProgressText,
                UnitIndex = checkpoint?.UnitIndex ?? 0,
                SubUnitIndex = checkpoint?.SubUnitIndex ?? 0,
                ErrorMessage = item.ErrorMessage,
                UpdatedAt = DateTime.Now
            });
        }
    }

    private Task WriteHistoryAsync(DocumentTranslationItem item, AppSettings settings, DateTime startedAt) =>
        historyService.AddAsync(new Models.TranslationHistoryRecord
        {
            Timestamp = DateTime.Now,
            SourcePath = item.SourcePath,
            OutputPath = item.OutputPath,
            ProviderType = settings.Ai.ProviderType,
            Model = settings.Ai.Model,
            Status = item.Status.ToString(),
            Progress = item.Progress,
            DurationSeconds = Math.Round((DateTime.Now - startedAt).TotalSeconds, 2),
            ErrorMessage = item.ErrorMessage
        });

    private static Task UpdateItemAsync(DocumentTranslationItem item, Action action)
    {
        if (Application.Current?.Dispatcher.CheckAccess() == true)
        {
            action();
            return Task.CompletedTask;
        }

        return Application.Current?.Dispatcher.InvokeAsync(action).Task ?? Task.Run(action);
    }
}
