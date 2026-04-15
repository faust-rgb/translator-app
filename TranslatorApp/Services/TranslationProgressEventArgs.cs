namespace TranslatorApp.Services;

public sealed class TranslationProgressEventArgs(string title, string partialText) : EventArgs
{
    public string Title { get; } = title;
    public string PartialText { get; } = partialText;
}
