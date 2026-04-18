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
    public int RetryCount { get; set; } = 2;
    public bool EnableStreaming { get; set; } = true;
    public string PromptTemplate { get; set; } =
        "你是一名专业文档翻译助手。请把下面内容从{source}翻译为{target}。仅返回译文，不要解释，不要添加标题，不要补充说明。请保留原有段落、列表、换行、数字、符号和占位符，不要漏译正文内容。";
}
