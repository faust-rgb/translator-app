using TranslatorApp.Models;

namespace TranslatorApp.Services;

public interface ITranslationHistoryService
{
    Task<IReadOnlyList<TranslationHistoryRecord>> LoadAsync();
    Task AddAsync(TranslationHistoryRecord record);
    Task ClearAsync();
}
