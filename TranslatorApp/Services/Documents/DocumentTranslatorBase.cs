using System.IO;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public abstract class DocumentTranslatorBase(ITextTranslationService textTranslationService, IAppLogService logService)
    : IDocumentTranslator
{
    public abstract bool CanHandle(string extension);

    public abstract Task TranslateAsync(TranslationJobContext context);

    protected async Task<string> TranslateBlockAsync(
        string text,
        string contextHint,
        AppSettings settings,
        PauseController pauseController,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        await pauseController.WaitIfPausedAsync(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return await textTranslationService.TranslateAsync(text, contextHint, settings, onPartialResponse, cancellationToken);
    }

    protected static string BuildOutputPath(string sourcePath, string outputDirectory)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var extension = Path.GetExtension(sourcePath);
        var directory = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.GetDirectoryName(sourcePath)!
            : outputDirectory;

        Directory.CreateDirectory(directory);
        return Path.Combine(directory, $"{fileName}.translated{extension}");
    }

    protected void Log(string message) => logService.Info(message);
}
