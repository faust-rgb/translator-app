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

    public async Task SaveAsync(AppSettings settings)
    {
        var directory = Path.GetDirectoryName(_settingsPath)!;
        Directory.CreateDirectory(directory);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        await File.WriteAllTextAsync(_settingsPath, json);
    }
}
