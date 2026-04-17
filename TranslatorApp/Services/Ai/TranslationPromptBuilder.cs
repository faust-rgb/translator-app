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
        $"文档上下文（仅供理解，不要翻译这一部分）：\n{request.ContextHint}\n\n" +
        "翻译要求：\n" +
        "1. 只翻译“待翻译文本”部分，仅返回译文。\n" +
        "2. 保留原文段落、换行、列表和占位符结构，不要把多个段落合并成一整段。\n" +
        "3. 不要遗漏任何可翻译英文；若当前片段是跨页续句，请结合上下文把当前片段完整译成中文。\n" +
        "4. 若待翻译文本是英文残句、断词续句或句中片段，也必须译成自然中文，不要原样照抄英文。\n" +
        "5. 除公认缩写、公式、代码、文献编号等必须保留的内容外，不要残留未翻译英文。\n" +
        (string.IsNullOrWhiteSpace(request.AdditionalRequirements)
            ? "\n"
            : $"补充要求：\n{request.AdditionalRequirements.Trim()}\n\n") +
        $"待翻译文本：\n<<<TEXT\n{request.Text}\nTEXT";
}
