using TranslatorApp.Models;

namespace TranslatorApp.Services;

public interface IRecoveryStateService
{
    Task<IReadOnlyList<DocumentCheckpoint>> LoadPendingAsync();
    Task<DocumentCheckpoint?> GetCheckpointAsync(string sourcePath);
    Task SaveCheckpointAsync(DocumentCheckpoint checkpoint);
    Task RemoveCheckpointAsync(string sourcePath);
}
