using System.IO;
using System.Text.Json;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class RecoveryStateService : IRecoveryStateService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
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
        var all = await LoadAllAsync();
        return all.FirstOrDefault(x => string.Equals(x.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase));
    }

    public async Task SaveCheckpointAsync(DocumentCheckpoint checkpoint)
    {
        var all = (await LoadAllAsync()).ToList();
        var index = all.FindIndex(x => string.Equals(x.SourcePath, checkpoint.SourcePath, StringComparison.OrdinalIgnoreCase));
        if (index >= 0)
        {
            all[index] = checkpoint;
        }
        else
        {
            all.Add(checkpoint);
        }

        await WriteAllAsync(all);
    }

    public async Task RemoveCheckpointAsync(string sourcePath)
    {
        var all = (await LoadAllAsync())
            .Where(x => !string.Equals(x.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
            .ToList();
        await WriteAllAsync(all);
    }

    private async Task<IReadOnlyList<DocumentCheckpoint>> LoadAllAsync()
    {
        if (!File.Exists(_path))
        {
            return Array.Empty<DocumentCheckpoint>();
        }

        var json = await File.ReadAllTextAsync(_path);
        return JsonSerializer.Deserialize<List<DocumentCheckpoint>>(json, JsonOptions) ?? [];
    }

    private async Task WriteAllAsync(IReadOnlyList<DocumentCheckpoint> checkpoints)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var json = JsonSerializer.Serialize(checkpoints, JsonOptions);
        await File.WriteAllTextAsync(_path, json);
    }
}
