using System.IO;
using System.Text.Json;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class TranslationHistoryService : ITranslationHistoryService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly SemaphoreSlim FileLock = new(1, 1);
    private readonly string _path = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TranslatorApp",
        "history.json");

    public async Task<IReadOnlyList<TranslationHistoryRecord>> LoadAsync()
    {
        await FileLock.WaitAsync();
        try
        {
            return await LoadUnsafeAsync();
        }
        finally
        {
            FileLock.Release();
        }
    }

    public async Task AddAsync(TranslationHistoryRecord record)
    {
        ArgumentNullException.ThrowIfNull(record);

        await FileLock.WaitAsync();
        try
        {
            var items = (await LoadUnsafeAsync()).ToList();
            items.Insert(0, record);
            if (items.Count > 200)
            {
                items = items.Take(200).ToList();
            }

            Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
            var json = JsonSerializer.Serialize(items, JsonOptions);
            await File.WriteAllTextAsync(_path, json);
        }
        finally
        {
            FileLock.Release();
        }
    }

    public async Task ClearAsync()
    {
        await FileLock.WaitAsync();
        try
        {
            if (File.Exists(_path))
            {
                File.Delete(_path);
            }
        }
        finally
        {
            FileLock.Release();
        }
    }

    private async Task<IReadOnlyList<TranslationHistoryRecord>> LoadUnsafeAsync()
    {
        if (!File.Exists(_path))
        {
            return Array.Empty<TranslationHistoryRecord>();
        }

        try
        {
            var json = await File.ReadAllTextAsync(_path);
            return JsonSerializer.Deserialize<List<TranslationHistoryRecord>>(json, JsonOptions) ?? [];
        }
        catch (JsonException)
        {
            return Array.Empty<TranslationHistoryRecord>();
        }
    }
}
