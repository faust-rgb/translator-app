using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Options;
using TranslatorApp.Configuration;

namespace TranslatorApp.Services;

public sealed class SettingsService(IOptions<AppSettings> initialOptions) : ISettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _settingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TranslatorApp",
        "user-settings.json");

    public Task<AppSettings> LoadAsync()
    {
        if (!File.Exists(_settingsPath))
        {
            return Task.FromResult(initialOptions.Value);
        }

        var json = File.ReadAllText(_settingsPath);
        var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
        return Task.FromResult(settings);
    }

    public async Task<AppSettings> LoadFromFileAsync(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("配置文件不存在。", path);
        }

        var json = await File.ReadAllTextAsync(path);
        return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
    }

    public async Task SaveAsync(AppSettings settings)
    {
        var directory = Path.GetDirectoryName(_settingsPath)!;
        Directory.CreateDirectory(directory);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        await File.WriteAllTextAsync(_settingsPath, json);
    }

    public async Task SaveToFileAsync(AppSettings settings, string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(settings, JsonOptions);
        await File.WriteAllTextAsync(path, json);
    }
}
