using TranslatorApp.Configuration;

namespace TranslatorApp.Services;

public interface ISettingsService
{
    Task<AppSettings> LoadAsync();
    Task<AppSettings> LoadFromFileAsync(string path);
    Task SaveAsync(AppSettings settings);
    Task SaveToFileAsync(AppSettings settings, string path);
}
