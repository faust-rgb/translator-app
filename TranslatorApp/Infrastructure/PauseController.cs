using System.Threading;

namespace TranslatorApp.Infrastructure;

public sealed class PauseController
{
    private readonly object _syncRoot = new();
    private TaskCompletionSource<bool> _resumeSignal = CreateSignal();
    private bool _isPaused;

    public bool IsPaused => _isPaused;

    public void Pause()
    {
        lock (_syncRoot)
        {
            if (_isPaused)
            {
                return;
            }

            _resumeSignal = CreateSignal();
            _isPaused = true;
        }
    }

    public void Resume()
    {
        TaskCompletionSource<bool>? signal = null;
        lock (_syncRoot)
        {
            if (!_isPaused)
            {
                return;
            }

            _isPaused = false;
            signal = _resumeSignal;
        }

        signal.TrySetResult(true);
    }

    public Task WaitIfPausedAsync(CancellationToken cancellationToken)
    {
        TaskCompletionSource<bool>? signal = null;
        lock (_syncRoot)
        {
            if (!_isPaused)
            {
                return Task.CompletedTask;
            }

            signal = _resumeSignal;
        }

        return signal.Task.WaitAsync(cancellationToken);
    }

    private static TaskCompletionSource<bool> CreateSignal() =>
        new(TaskCreationOptions.RunContinuationsAsynchronously);
}
