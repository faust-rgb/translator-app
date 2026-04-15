using TranslatorApp.Models;

namespace TranslatorApp.Services;

public interface IGlossaryService
{
    Task<IReadOnlyList<GlossaryEntry>> LoadAsync(string glossaryPath, CancellationToken cancellationToken);
    string BuildPromptSection(string text, IReadOnlyList<GlossaryEntry> entries);
}
