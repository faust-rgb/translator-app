using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Options;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;

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

        return NormalizeSettings(settings);
    }

    public async Task<AppSettings> LoadFromFileAsync(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

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

        return NormalizeSettings(settings);
    }

    public async Task SaveAsync(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        settings = NormalizeSettings(settings);

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
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        settings = NormalizeSettings(settings);
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
        ArgumentNullException.ThrowIfNull(settings);

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
                ServerErrorRetryCount = settings.Translation.ServerErrorRetryCount,
                TimeoutRetryCount = settings.Translation.TimeoutRetryCount,
                RetryCount = settings.Translation.RetryCount,
                PdfColumnGapRatio = settings.Translation.PdfColumnGapRatio,
                PdfColumnMinWordsPerSide = settings.Translation.PdfColumnMinWordsPerSide,
                PdfColumnMinVerticalSpanRatio = settings.Translation.PdfColumnMinVerticalSpanRatio,
                PdfMarginNoiseSideRatio = settings.Translation.PdfMarginNoiseSideRatio,
                PdfMarginNoiseVerticalAspectRatio = settings.Translation.PdfMarginNoiseVerticalAspectRatio,
                PdfMarginNoiseShortTokenLength = settings.Translation.PdfMarginNoiseShortTokenLength,
                PdfParagraphGroupingMaxVerticalGapRatio = settings.Translation.PdfParagraphGroupingMaxVerticalGapRatio,
                PdfContinuationMergeMaxVerticalGapRatio = settings.Translation.PdfContinuationMergeMaxVerticalGapRatio,
                PdfLineMergeMaxVerticalGapRatio = settings.Translation.PdfLineMergeMaxVerticalGapRatio,
                PdfParagraphLeftAlignToleranceRatio = settings.Translation.PdfParagraphLeftAlignToleranceRatio,
                PdfParagraphRightAlignToleranceRatio = settings.Translation.PdfParagraphRightAlignToleranceRatio,
                PdfParagraphOverlapThreshold = settings.Translation.PdfParagraphOverlapThreshold,
                PdfParagraphHorizontalGapRatio = settings.Translation.PdfParagraphHorizontalGapRatio,
                PdfParagraphRangeRelationRatio = settings.Translation.PdfParagraphRangeRelationRatio,
                PdfParagraphMinWidthRatio = settings.Translation.PdfParagraphMinWidthRatio,
                PdfParagraphLooseWrapForwardRatio = settings.Translation.PdfParagraphLooseWrapForwardRatio,
                PdfParagraphLooseWrapBackwardRatio = settings.Translation.PdfParagraphLooseWrapBackwardRatio,
                PromptTemplate = settings.Translation.PromptTemplate
            },
            Ocr = new OcrSettings
            {
                EnableOcrForScannedPdf = settings.Ocr.EnableOcrForScannedPdf,
                Language = settings.Ocr.Language,
                RenderScale = settings.Ocr.RenderScale,
                MinimumNativeTextWords = settings.Ocr.MinimumNativeTextWords,
                SparseTextCoverageThreshold = settings.Ocr.SparseTextCoverageThreshold,
                SparseTextBlockThreshold = settings.Ocr.SparseTextBlockThreshold,
                MinimumAcceptedConfidence = settings.Ocr.MinimumAcceptedConfidence,
                OcrBlockMergeVerticalGapRatio = settings.Ocr.OcrBlockMergeVerticalGapRatio,
                OcrBlockMergeIndentationRatio = settings.Ocr.OcrBlockMergeIndentationRatio,
                OcrBlockMergeWidthDifferenceRatio = settings.Ocr.OcrBlockMergeWidthDifferenceRatio,
                OcrColumnCenterToleranceRatio = settings.Ocr.OcrColumnCenterToleranceRatio,
                OcrColumnLeftToleranceRatio = settings.Ocr.OcrColumnLeftToleranceRatio,
                OcrColumnOverlapThreshold = settings.Ocr.OcrColumnOverlapThreshold
            }
        };
    }

    private static AppSettings NormalizeSettings(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        settings.Ai ??= new AiSettings();
        settings.Translation ??= new TranslationSettings();
        settings.Ocr ??= new OcrSettings();

        settings.Ai.ProviderType = string.IsNullOrWhiteSpace(settings.Ai.ProviderType)
            ? "AnthropicCompatible"
            : settings.Ai.ProviderType.Trim();
        settings.Ai.BaseUrl = string.IsNullOrWhiteSpace(settings.Ai.BaseUrl)
            ? "https://api.anthropic.com"
            : settings.Ai.BaseUrl.Trim();
        settings.Ai.Model = string.IsNullOrWhiteSpace(settings.Ai.Model)
            ? "claude-3-7-sonnet-latest"
            : settings.Ai.Model.Trim();
        settings.Ai.AnthropicVersion = string.IsNullOrWhiteSpace(settings.Ai.AnthropicVersion)
            ? "2023-06-01"
            : settings.Ai.AnthropicVersion.Trim();
        settings.Ai.TimeoutSeconds = Math.Clamp(settings.Ai.TimeoutSeconds, 5, 3600);
        settings.Ai.MaxTokens = Math.Clamp(settings.Ai.MaxTokens, 1, 128000);
        settings.Ai.Temperature = Math.Clamp(settings.Ai.Temperature, 0, 2);

        settings.Translation.SourceLanguage = string.IsNullOrWhiteSpace(settings.Translation.SourceLanguage)
            ? "自动检测"
            : settings.Translation.SourceLanguage.Trim();
        settings.Translation.TargetLanguage = string.IsNullOrWhiteSpace(settings.Translation.TargetLanguage)
            ? "中文"
            : settings.Translation.TargetLanguage.Trim();
        settings.Translation.EbookOutputFormat = string.Equals(settings.Translation.EbookOutputFormat, "DOCX", StringComparison.OrdinalIgnoreCase)
            ? "DOCX"
            : "EPUB";
        settings.Translation.OutputFontFamily = string.IsNullOrWhiteSpace(settings.Translation.OutputFontFamily)
            ? PdfSharpFontResolver.DefaultFontFamily
            : settings.Translation.OutputFontFamily.Trim();
        settings.Translation.OutputFontSize = Math.Clamp(settings.Translation.OutputFontSize, 6, 72);
        settings.Translation.RangeStart = Math.Max(1, settings.Translation.RangeStart);
        settings.Translation.RangeEnd = Math.Max(0, settings.Translation.RangeEnd);
        settings.Translation.MaxParallelDocuments = Math.Max(1, settings.Translation.MaxParallelDocuments);
        settings.Translation.MaxParallelBlocks = Math.Max(1, settings.Translation.MaxParallelBlocks);
        settings.Translation.MaxGlobalTranslationRequests = Math.Max(1, settings.Translation.MaxGlobalTranslationRequests);
        settings.Translation.ServerErrorRetryCount = Math.Max(0, settings.Translation.ServerErrorRetryCount);
        settings.Translation.TimeoutRetryCount = Math.Max(0, settings.Translation.TimeoutRetryCount);
        settings.Translation.RetryCount = Math.Max(0, settings.Translation.RetryCount);
        settings.Translation.PdfColumnGapRatio = Math.Clamp(settings.Translation.PdfColumnGapRatio, 0.005, 0.25);
        settings.Translation.PdfColumnMinWordsPerSide = Math.Clamp(settings.Translation.PdfColumnMinWordsPerSide, 1, 20);
        settings.Translation.PdfColumnMinVerticalSpanRatio = Math.Clamp(settings.Translation.PdfColumnMinVerticalSpanRatio, 0.05, 0.9);
        settings.Translation.PdfMarginNoiseSideRatio = Math.Clamp(settings.Translation.PdfMarginNoiseSideRatio, 0.01, 0.4);
        settings.Translation.PdfMarginNoiseVerticalAspectRatio = Math.Clamp(settings.Translation.PdfMarginNoiseVerticalAspectRatio, 0.5, 5.0);
        settings.Translation.PdfMarginNoiseShortTokenLength = Math.Clamp(settings.Translation.PdfMarginNoiseShortTokenLength, 1, 20);
        settings.Translation.PdfParagraphGroupingMaxVerticalGapRatio = Math.Clamp(settings.Translation.PdfParagraphGroupingMaxVerticalGapRatio, 0.5, 8.0);
        settings.Translation.PdfContinuationMergeMaxVerticalGapRatio = Math.Clamp(settings.Translation.PdfContinuationMergeMaxVerticalGapRatio, 0.5, 5.0);
        settings.Translation.PdfLineMergeMaxVerticalGapRatio = Math.Clamp(settings.Translation.PdfLineMergeMaxVerticalGapRatio, 0.5, 4.0);
        settings.Translation.PdfParagraphLeftAlignToleranceRatio = Math.Clamp(settings.Translation.PdfParagraphLeftAlignToleranceRatio, 0.2, 8.0);
        settings.Translation.PdfParagraphRightAlignToleranceRatio = Math.Clamp(settings.Translation.PdfParagraphRightAlignToleranceRatio, 0.2, 8.0);
        settings.Translation.PdfParagraphOverlapThreshold = Math.Clamp(settings.Translation.PdfParagraphOverlapThreshold, 0.01, 0.99);
        settings.Translation.PdfParagraphHorizontalGapRatio = Math.Clamp(settings.Translation.PdfParagraphHorizontalGapRatio, 0.5, 12.0);
        settings.Translation.PdfParagraphRangeRelationRatio = Math.Clamp(settings.Translation.PdfParagraphRangeRelationRatio, 0.5, 12.0);
        settings.Translation.PdfParagraphMinWidthRatio = Math.Clamp(settings.Translation.PdfParagraphMinWidthRatio, 1.0, 30.0);
        settings.Translation.PdfParagraphLooseWrapForwardRatio = Math.Clamp(settings.Translation.PdfParagraphLooseWrapForwardRatio, 0.5, 12.0);
        settings.Translation.PdfParagraphLooseWrapBackwardRatio = Math.Clamp(settings.Translation.PdfParagraphLooseWrapBackwardRatio, 0.5, 12.0);
        settings.Translation.PromptTemplate = string.IsNullOrWhiteSpace(settings.Translation.PromptTemplate)
            ? "你是一名专业文档翻译助手。请把下面内容从{source}翻译为{target}。仅返回译文，不要解释，不要添加标题，不要补充说明。请保留原有段落、列表、换行、数字、符号和占位符，不要漏译正文内容。"
            : settings.Translation.PromptTemplate.Trim();

        settings.Ocr.Language = string.IsNullOrWhiteSpace(settings.Ocr.Language)
            ? "chi_sim+eng"
            : settings.Ocr.Language.Trim();
        settings.Ocr.RenderScale = Math.Clamp(settings.Ocr.RenderScale, 0.5, 3.0);
        settings.Ocr.MinimumNativeTextWords = Math.Clamp(settings.Ocr.MinimumNativeTextWords, 0, 500);
        settings.Ocr.SparseTextCoverageThreshold = Math.Clamp(settings.Ocr.SparseTextCoverageThreshold, 0.001, 0.5);
        settings.Ocr.SparseTextBlockThreshold = Math.Clamp(settings.Ocr.SparseTextBlockThreshold, 1, 50);
        settings.Ocr.MinimumAcceptedConfidence = Math.Clamp(settings.Ocr.MinimumAcceptedConfidence, 0, 100);
        settings.Ocr.OcrBlockMergeVerticalGapRatio = Math.Clamp(settings.Ocr.OcrBlockMergeVerticalGapRatio, 0.2, 3.0);
        settings.Ocr.OcrBlockMergeIndentationRatio = Math.Clamp(settings.Ocr.OcrBlockMergeIndentationRatio, 0.2, 5.0);
        settings.Ocr.OcrBlockMergeWidthDifferenceRatio = Math.Clamp(settings.Ocr.OcrBlockMergeWidthDifferenceRatio, 0.1, 2.0);
        settings.Ocr.OcrColumnCenterToleranceRatio = Math.Clamp(settings.Ocr.OcrColumnCenterToleranceRatio, 0.01, 0.5);
        settings.Ocr.OcrColumnLeftToleranceRatio = Math.Clamp(settings.Ocr.OcrColumnLeftToleranceRatio, 0.01, 0.5);
        settings.Ocr.OcrColumnOverlapThreshold = Math.Clamp(settings.Ocr.OcrColumnOverlapThreshold, 0.05, 0.95);
        return settings;
    }
}
