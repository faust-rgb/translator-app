using TranslatorApp.Infrastructure;
using TranslatorApp.Configuration;

namespace TranslatorApp.Models;

public sealed class TranslationJobContext
{
    public required AppSettings Settings { get; init; }
    public required DocumentTranslationItem Item { get; init; }
    public required CancellationToken CancellationToken { get; init; }
    public required PauseController PauseController { get; init; }
    public required Func<int, string, Task> ReportProgressAsync { get; init; }
    public required Func<int, int, string, Task> SaveCheckpointAsync { get; init; }
    public int ResumeUnitIndex { get; init; }
    public int ResumeSubUnitIndex { get; init; }
}
