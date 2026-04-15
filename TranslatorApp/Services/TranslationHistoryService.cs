using System.IO;
using System.Text.Json;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class TranslationHistoryService : ITranslationHistoryService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TranslatorApp",
        "history.json");

    public async Task<IReadOnlyList<TranslationHistoryRecord>> LoadAsync()
    {
        if (!File.Exists(_path))
        {
            return Array.Empty<TranslationHistoryRecord>();
        }

        var json = await File.ReadAllTextAsync(_path);
        return JsonSerializer.Deserialize<List<TranslationHistoryRecord>>(json, JsonOptions) ?? [];
    }

    public async Task AddAsync(TranslationHistoryRecord record)
    {
        var items = (await LoadAsync()).ToList();
        items.Insert(0, record);
        if (items.Count > 200)
        {
            items = items.Take(200).ToList();
        }

        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var json = JsonSerializer.Serialize(items, JsonOptions);
        await File.WriteAllTextAsync(_path, json);
    }

    public Task ClearAsync()
    {
        if (File.Exists(_path))
        {
            File.Delete(_path);
        }

        return Task.CompletedTask;
    }
}
