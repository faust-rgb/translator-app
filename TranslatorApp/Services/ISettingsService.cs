using TranslatorApp.Configuration;

namespace TranslatorApp.Services;

public interface ISettingsService
{
    Task<AppSettings> LoadAsync();
    Task SaveAsync(AppSettings settings);
}
