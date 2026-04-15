using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public interface IDocumentTranslator
{
    bool CanHandle(string extension);
    Task TranslateAsync(TranslationJobContext context);
}
