using System.IO;
using System.Text.Json;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class RecoveryStateService : IRecoveryStateService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly SemaphoreSlim FileLock = new(1, 1);
    private readonly string _path = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TranslatorApp",
        "recovery-state.json");

    public async Task<IReadOnlyList<DocumentCheckpoint>> LoadPendingAsync()
    {
        var all = await LoadAllAsync();
        return all
            .Where(x => string.Equals(x.Status, nameof(DocumentStatus.Running), StringComparison.OrdinalIgnoreCase)
                     || string.Equals(x.Status, nameof(DocumentStatus.Paused), StringComparison.OrdinalIgnoreCase)
                     || string.Equals(x.Status, nameof(DocumentStatus.Pending), StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(x => x.UpdatedAt)
            .ToList();
    }

    public async Task<DocumentCheckpoint?> GetCheckpointAsync(string sourcePath)
    {
        await FileLock.WaitAsync();
        try
        {
            var all = await LoadAllUnsafeAsync();
            return all.FirstOrDefault(x => string.Equals(x.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            FileLock.Release();
        }
    }

    public async Task SaveCheckpointAsync(DocumentCheckpoint checkpoint)
    {
        await FileLock.WaitAsync();
        try
        {
            var all = (await LoadAllUnsafeAsync()).ToList();
            var index = all.FindIndex(x => string.Equals(x.SourcePath, checkpoint.SourcePath, StringComparison.OrdinalIgnoreCase));
            if (index >= 0)
            {
                all[index] = checkpoint;
            }
            else
            {
                all.Add(checkpoint);
            }

            await WriteAllUnsafeAsync(all);
        }
        finally
        {
            FileLock.Release();
        }
    }

    public async Task RemoveCheckpointAsync(string sourcePath)
    {
        await FileLock.WaitAsync();
        try
        {
            var all = (await LoadAllUnsafeAsync())
                .Where(x => !string.Equals(x.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
                .ToList();
            await WriteAllUnsafeAsync(all);
        }
        finally
        {
            FileLock.Release();
        }
    }

    private async Task<IReadOnlyList<DocumentCheckpoint>> LoadAllAsync()
    {
        await FileLock.WaitAsync();
        try
        {
            return await LoadAllUnsafeAsync();
        }
        finally
        {
            FileLock.Release();
        }
    }

    private async Task<IReadOnlyList<DocumentCheckpoint>> LoadAllUnsafeAsync()
    {
        if (!File.Exists(_path))
        {
            return Array.Empty<DocumentCheckpoint>();
        }

        try
        {
            var json = await File.ReadAllTextAsync(_path);
            return JsonSerializer.Deserialize<List<DocumentCheckpoint>>(json, JsonOptions) ?? [];
        }
        catch (JsonException)
        {
            return Array.Empty<DocumentCheckpoint>();
        }
    }

    private async Task WriteAllUnsafeAsync(IReadOnlyList<DocumentCheckpoint> checkpoints)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var json = JsonSerializer.Serialize(checkpoints, JsonOptions);
        await File.WriteAllTextAsync(_path, json);
    }
}
