using TranslatorApp.Configuration;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Ai;

public sealed class TranslationPromptBuilder : ITranslationPromptBuilder
{
    public string BuildSystemPrompt(TranslationRequest request) =>
        $"{request.PromptTemplate}\n{request.GlossaryPrompt}".Trim()
            .Replace("{source}", request.SourceLanguage, StringComparison.OrdinalIgnoreCase)
            .Replace("{target}", request.TargetLanguage, StringComparison.OrdinalIgnoreCase);

    public string BuildUserPrompt(TranslationRequest request) =>
        $"文档上下文：{request.ContextHint}\n\n待翻译文本：\n{request.Text}";
}
