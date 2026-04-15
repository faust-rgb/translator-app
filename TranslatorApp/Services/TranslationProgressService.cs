using System.Windows;

namespace TranslatorApp.Services;

public sealed class TranslationProgressService : ITranslationProgressService
{
    public event EventHandler<TranslationProgressEventArgs>? StreamUpdated;

    public void Publish(string title, string partialText)
    {
        var args = new TranslationProgressEventArgs(title, partialText);
        if (Application.Current?.Dispatcher.CheckAccess() == true)
        {
            StreamUpdated?.Invoke(this, args);
            return;
        }

        Application.Current?.Dispatcher.Invoke(() => StreamUpdated?.Invoke(this, args));
    }

    public void Clear() => Publish(string.Empty, string.Empty);
}
