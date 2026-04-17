using System.Windows;
using System.Windows.Threading;

namespace TranslatorApp.Services;

public sealed class TranslationProgressService : ITranslationProgressService
{
    private static readonly TimeSpan PublishInterval = TimeSpan.FromMilliseconds(150);
    private readonly object _syncRoot = new();
    private TranslationProgressEventArgs? _pendingArgs;
    private bool _isFlushScheduled;

    public event EventHandler<TranslationProgressEventArgs>? StreamUpdated;

    public void Publish(string title, string partialText)
    {
        lock (_syncRoot)
        {
            _pendingArgs = new TranslationProgressEventArgs(title, partialText);
            if (_isFlushScheduled)
            {
                return;
            }

            _isFlushScheduled = true;
        }

        _ = FlushPendingAsync();
    }

    public void Clear() => Publish(string.Empty, string.Empty);

    private async Task FlushPendingAsync()
    {
        while (true)
        {
            TranslationProgressEventArgs? args;
            lock (_syncRoot)
            {
                args = _pendingArgs;
                _pendingArgs = null;
            }

            if (args is not null)
            {
                await DispatchAsync(args);
            }

            await Task.Delay(PublishInterval);

            lock (_syncRoot)
            {
                if (_pendingArgs is null)
                {
                    _isFlushScheduled = false;
                    return;
                }
            }
        }
    }

    private Task DispatchAsync(TranslationProgressEventArgs args)
    {
        if (Application.Current?.Dispatcher.CheckAccess() == true)
        {
            StreamUpdated?.Invoke(this, args);
            return Task.CompletedTask;
        }

        if (Application.Current?.Dispatcher is Dispatcher dispatcher)
        {
            return dispatcher.InvokeAsync(() => StreamUpdated?.Invoke(this, args)).Task;
        }

        StreamUpdated?.Invoke(this, args);
        return Task.CompletedTask;
    }
}
