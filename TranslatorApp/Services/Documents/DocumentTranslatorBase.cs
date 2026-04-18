using System.IO;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public abstract class DocumentTranslatorBase(ITextTranslationService textTranslationService, IAppLogService logService)
    : IDocumentTranslator
{
    public abstract bool CanHandle(string extension);

    public abstract Task TranslateAsync(TranslationJobContext context);

    protected async Task<string> TranslateBlockAsync(
        string text,
        string contextHint,
        AppSettings settings,
        PauseController pauseController,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        await pauseController.WaitIfPausedAsync(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return await textTranslationService.TranslateAsync(text, contextHint, settings, onPartialResponse, cancellationToken);
    }

    protected async Task<IReadOnlyList<string>> TranslateBatchAsync(
        IReadOnlyList<TranslationBlock> blocks,
        AppSettings settings,
        PauseController pauseController,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        if (blocks.Count == 0)
        {
            return Array.Empty<string>();
        }

        var results = new string[blocks.Count];
        var maxConcurrency = Math.Min(blocks.Count, GetBlockTranslationConcurrency(settings));
        if (maxConcurrency <= 1)
        {
            for (var i = 0; i < blocks.Count; i++)
            {
                results[i] = await TranslateBlockAsync(
                    blocks[i].Text,
                    blocks[i].ContextHint,
                    settings,
                    pauseController,
                    onPartialResponse,
                    cancellationToken);
            }

            return results;
        }

        using var semaphore = new SemaphoreSlim(maxConcurrency);
        var previewGate = new object();
        var previewAssigned = false;

        var tasks = blocks.Select(async (block, index) =>
        {
            await semaphore.WaitAsync(cancellationToken);
            try
            {
                Func<string, Task>? partialCallback = null;
                if (onPartialResponse is not null)
                {
                    lock (previewGate)
                    {
                        if (!previewAssigned)
                        {
                            previewAssigned = true;
                            partialCallback = onPartialResponse;
                        }
                    }
                }

                results[index] = await TranslateBlockAsync(
                    block.Text,
                    block.ContextHint,
                    settings,
                    pauseController,
                    partialCallback,
                    cancellationToken);
            }
            finally
            {
                semaphore.Release();
            }
        });

        await Task.WhenAll(tasks);
        return results;
    }

    protected static string BuildOutputPath(string sourcePath, string outputDirectory, string? outputExtension = null)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var extension = string.IsNullOrWhiteSpace(outputExtension)
            ? Path.GetExtension(sourcePath)
            : outputExtension;
        var directory = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.GetDirectoryName(sourcePath)!
            : outputDirectory;

        Directory.CreateDirectory(directory);
        return Path.Combine(directory, $"{fileName}.translated{extension}");
    }

    protected void Log(string message) => logService.Info(message);

    protected static int GetBlockTranslationConcurrency(AppSettings settings) =>
        Math.Max(1, settings.Translation.MaxParallelBlocks);

    protected static (int Start, int End) GetRequestedRange(AppSettings settings, int totalUnits)
    {
        totalUnits = Math.Max(0, totalUnits);
        if (totalUnits == 0)
        {
            return (1, 0);
        }

        var start = Math.Max(1, settings.Translation.RangeStart);
        var end = settings.Translation.RangeEnd <= 0 ? totalUnits : Math.Max(start, settings.Translation.RangeEnd);
        start = Math.Min(start, totalUnits);
        end = Math.Min(end, totalUnits);
        return (start, end);
    }

    protected static bool IsWithinRequestedRange(int oneBasedIndex, (int Start, int End) range) =>
        range.End >= range.Start && oneBasedIndex >= range.Start && oneBasedIndex <= range.End;

    protected static string DescribeRequestedRange(string unitName, (int Start, int End) range) =>
        range.End < range.Start ? $"未命中任何{unitName}" : $"{unitName} {range.Start}-{range.End}";

    protected readonly record struct TranslationBlock(string Text, string ContextHint);
}
