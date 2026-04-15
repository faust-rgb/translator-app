using System.Threading;

namespace TranslatorApp.Infrastructure;

public sealed class PauseController
{
    private volatile TaskCompletionSource<bool> _resumeSignal = CreateSignal();

    public bool IsPaused { get; private set; }

    public void Pause()
    {
        if (IsPaused)
        {
            return;
        }

        IsPaused = true;
        _resumeSignal = CreateSignal();
    }

    public void Resume()
    {
        if (!IsPaused)
        {
            return;
        }

        IsPaused = false;
        _resumeSignal.TrySetResult(true);
    }

    public Task WaitIfPausedAsync(CancellationToken cancellationToken)
    {
        if (!IsPaused)
        {
            return Task.CompletedTask;
        }

        return _resumeSignal.Task.WaitAsync(cancellationToken);
    }

    private static TaskCompletionSource<bool> CreateSignal() =>
        new(TaskCreationOptions.RunContinuationsAsynchronously);
}
