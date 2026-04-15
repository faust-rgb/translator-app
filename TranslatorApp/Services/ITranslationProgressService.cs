namespace TranslatorApp.Services;

public interface ITranslationProgressService
{
    event EventHandler<TranslationProgressEventArgs>? StreamUpdated;
    void Publish(string title, string partialText);
    void Clear();
}
