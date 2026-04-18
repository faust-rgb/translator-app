using TranslatorApp.Infrastructure;

namespace TranslatorApp.Configuration;

public sealed class TranslationSettings
{
    public string SourceLanguage { get; set; } = "自动检测";
    public string TargetLanguage { get; set; } = "中文";
    public string OutputDirectory { get; set; } = string.Empty;
    public int RangeStart { get; set; } = 1;
    public int RangeEnd { get; set; } = 0;
    public string EbookOutputFormat { get; set; } = "EPUB";
    public string CalibreExecutablePath { get; set; } = string.Empty;
    public string OutputFontFamily { get; set; } = PdfSharpFontResolver.DefaultFontFamily;
    public double OutputFontSize { get; set; } = 11;
    public int MaxParallelDocuments { get; set; } = 1;
    public int MaxParallelBlocks { get; set; } = 1;
    public int MaxGlobalTranslationRequests { get; set; } = 1;
    public string GlossaryPath { get; set; } = string.Empty;
    public bool ExportBilingualDocument { get; set; } = true;
    public int ServerErrorRetryCount { get; set; } = 2;
    public int TimeoutRetryCount { get; set; } = 1;
    public bool EnableStreaming { get; set; } = true;
    public double PdfColumnGapRatio { get; set; } = 0.04;
    public int PdfColumnMinWordsPerSide { get; set; } = 3;
    public double PdfColumnMinVerticalSpanRatio { get; set; } = 0.25;
    public double PdfMarginNoiseSideRatio { get; set; } = 0.1;
    public double PdfMarginNoiseVerticalAspectRatio { get; set; } = 1.4;
    public int PdfMarginNoiseShortTokenLength { get; set; } = 3;
    public double PdfParagraphGroupingMaxVerticalGapRatio { get; set; } = 3.2;
    public double PdfContinuationMergeMaxVerticalGapRatio { get; set; } = 1.8;
    public double PdfLineMergeMaxVerticalGapRatio { get; set; } = 1.35;
    public double PdfParagraphLeftAlignToleranceRatio { get; set; } = 1.6;
    public double PdfParagraphRightAlignToleranceRatio { get; set; } = 2.4;
    public double PdfParagraphOverlapThreshold { get; set; } = 0.3;
    public double PdfParagraphHorizontalGapRatio { get; set; } = 4.5;
    public double PdfParagraphRangeRelationRatio { get; set; } = 4.0;
    public double PdfParagraphMinWidthRatio { get; set; } = 10.0;
    public double PdfParagraphLooseWrapForwardRatio { get; set; } = 6.5;
    public double PdfParagraphLooseWrapBackwardRatio { get; set; } = 5.0;
    public string PromptTemplate { get; set; } =
        "你是一名专业文档翻译助手。请把下面内容从{source}翻译为{target}。仅返回译文，不要解释，不要添加标题，不要补充说明。请保留原有段落、列表、换行、数字、符号和占位符，不要漏译正文内容。";

    // 兼容旧配置字段：如果新字段缺失，仍可回退到旧的统一重试次数。
    public int RetryCount { get; set; } = 2;
}
