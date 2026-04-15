using TranslatorApp.Configuration;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Ai;

public interface ITranslationPromptBuilder
{
    string BuildSystemPrompt(TranslationRequest request);
    string BuildUserPrompt(TranslationRequest request);
}
