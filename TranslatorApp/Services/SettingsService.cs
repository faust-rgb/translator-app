using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Options;
using TranslatorApp.Configuration;

namespace TranslatorApp.Services;

public sealed class SettingsService(
    IOptions<AppSettings> initialOptions,
    ISecureApiKeyStorage secureApiKeyStorage) : ISettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _settingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TranslatorApp",
        "user-settings.json");

    public async Task<AppSettings> LoadAsync()
    {
        AppSettings settings;

        if (!File.Exists(_settingsPath))
        {
            settings = initialOptions.Value;
        }
        else
        {
            var json = await File.ReadAllTextAsync(_settingsPath);
            settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
        }

        // 从安全存储加载 API Key
        var secureKey = await secureApiKeyStorage.GetAsync(settings.Ai.ProviderType);
        if (!string.IsNullOrEmpty(secureKey))
        {
            settings.Ai.ApiKey = secureKey;
        }

        return settings;
    }

    public async Task<AppSettings> LoadFromFileAsync(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("配置文件不存在。", path);
        }

        var json = await File.ReadAllTextAsync(path);
        var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();

        // 优先使用配置文件中的 API Key；仅当文件未提供时才回退到安全存储。
        if (string.IsNullOrWhiteSpace(settings.Ai.ApiKey))
        {
            var secureKey = await secureApiKeyStorage.GetAsync(settings.Ai.ProviderType);
            if (!string.IsNullOrEmpty(secureKey))
            {
                settings.Ai.ApiKey = secureKey;
            }
        }

        return settings;
    }

    public async Task SaveAsync(AppSettings settings)
    {
        // 保存 API Key 到安全存储
        if (!string.IsNullOrEmpty(settings.Ai.ApiKey))
        {
            await secureApiKeyStorage.SetAsync(settings.Ai.ProviderType, settings.Ai.ApiKey);
        }

        // 创建不包含敏感信息的副本
        var settingsForSave = CreateSafeSettingsCopy(settings);

        var directory = Path.GetDirectoryName(_settingsPath)!;
        Directory.CreateDirectory(directory);
        var json = JsonSerializer.Serialize(settingsForSave, JsonOptions);
        await File.WriteAllTextAsync(_settingsPath, json);
    }

    public async Task SaveToFileAsync(AppSettings settings, string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // 导出配置文件时保留 API Key，方便跨机器/跨目录迁移。
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        await File.WriteAllTextAsync(path, json);
    }

    /// <summary>
    /// 创建不包含敏感信息的设置副本。
    /// </summary>
    private static AppSettings CreateSafeSettingsCopy(AppSettings settings)
    {
        return new AppSettings
        {
            Ai = new AiSettings
            {
                ProviderType = settings.Ai.ProviderType,
                BaseUrl = settings.Ai.BaseUrl,
                ApiKey = string.Empty, // 不保存 API Key 到普通配置文件
                Model = settings.Ai.Model,
                AnthropicVersion = settings.Ai.AnthropicVersion,
                Temperature = settings.Ai.Temperature,
                MaxTokens = settings.Ai.MaxTokens,
                CustomHeaders = settings.Ai.CustomHeaders,
                TimeoutSeconds = settings.Ai.TimeoutSeconds
            },
            Translation = new TranslationSettings
            {
                SourceLanguage = settings.Translation.SourceLanguage,
                TargetLanguage = settings.Translation.TargetLanguage,
                OutputDirectory = settings.Translation.OutputDirectory,
                RangeStart = settings.Translation.RangeStart,
                RangeEnd = settings.Translation.RangeEnd,
                EbookOutputFormat = settings.Translation.EbookOutputFormat,
                CalibreExecutablePath = settings.Translation.CalibreExecutablePath,
                OutputFontFamily = settings.Translation.OutputFontFamily,
                OutputFontSize = settings.Translation.OutputFontSize,
                MaxParallelDocuments = settings.Translation.MaxParallelDocuments,
                MaxParallelBlocks = settings.Translation.MaxParallelBlocks,
                MaxGlobalTranslationRequests = settings.Translation.MaxGlobalTranslationRequests,
                GlossaryPath = settings.Translation.GlossaryPath,
                ExportBilingualDocument = settings.Translation.ExportBilingualDocument,
                EnableStreaming = settings.Translation.EnableStreaming,
                RetryCount = settings.Translation.RetryCount,
                PromptTemplate = settings.Translation.PromptTemplate
            },
            Ocr = new OcrSettings
            {
                EnableOcrForScannedPdf = settings.Ocr.EnableOcrForScannedPdf,
                Language = settings.Ocr.Language,
                RenderScale = settings.Ocr.RenderScale,
                MinimumNativeTextWords = settings.Ocr.MinimumNativeTextWords
            }
        };
    }
}
