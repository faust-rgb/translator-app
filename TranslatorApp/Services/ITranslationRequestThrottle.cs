namespace TranslatorApp.Services;

public interface ITranslationRequestThrottle
{
    Task<IAsyncDisposable> AcquireAsync(int maxConcurrency, CancellationToken cancellationToken);
}
