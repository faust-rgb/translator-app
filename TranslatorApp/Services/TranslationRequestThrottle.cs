namespace TranslatorApp.Services;

public sealed class TranslationRequestThrottle : ITranslationRequestThrottle
{
    private readonly object _syncRoot = new();
    private readonly LinkedList<Waiter> _waiters = new();
    private int _currentConcurrency;
    private int _maxConcurrency = 1;

    public Task<IAsyncDisposable> AcquireAsync(int maxConcurrency, CancellationToken cancellationToken)
    {
        if (maxConcurrency <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maxConcurrency), maxConcurrency, "并发数必须大于 0。");
        }

        lock (_syncRoot)
        {
            _maxConcurrency = maxConcurrency;
            if (_currentConcurrency < _maxConcurrency)
            {
                _currentConcurrency++;
                return Task.FromResult<IAsyncDisposable>(new Releaser(this));
            }

            var waiter = new Waiter();
            waiter.Node = _waiters.AddLast(waiter);
            if (cancellationToken.CanBeCanceled)
            {
                waiter.CancellationRegistration = cancellationToken.Register(static state =>
                {
                    var target = (Waiter)state!;
                    target.Cancel();
                }, waiter);
            }

            return waiter.Task;
        }
    }

    private void Release()
    {
        Waiter? nextWaiter = null;

        lock (_syncRoot)
        {
            _currentConcurrency = Math.Max(0, _currentConcurrency - 1);
            while (_currentConcurrency < _maxConcurrency && _waiters.Count > 0)
            {
                var node = _waiters.First!;
                _waiters.RemoveFirst();
                var waiter = node.Value;
                if (waiter.TryActivate())
                {
                    _currentConcurrency++;
                    nextWaiter = waiter;
                    break;
                }
            }
        }

        nextWaiter?.SetResult(new Releaser(this));
    }

    private sealed class Waiter
    {
        private readonly TaskCompletionSource<IAsyncDisposable> _tcs = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _state;

        public LinkedListNode<Waiter>? Node { get; set; }
        public CancellationTokenRegistration CancellationRegistration { get; set; }
        public Task<IAsyncDisposable> Task => _tcs.Task;

        public bool TryActivate() => Interlocked.CompareExchange(ref _state, 1, 0) == 0;

        public void SetResult(IAsyncDisposable releaser)
        {
            CancellationRegistration.Dispose();
            _tcs.TrySetResult(releaser);
        }

        public void Cancel()
        {
            if (Interlocked.CompareExchange(ref _state, 2, 0) != 0)
            {
                return;
            }

            CancellationRegistration.Dispose();
            _tcs.TrySetCanceled();
        }
    }

    private sealed class Releaser(TranslationRequestThrottle owner) : IAsyncDisposable
    {
        private int _disposed;

        public ValueTask DisposeAsync()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
            {
                owner.Release();
            }

            return ValueTask.CompletedTask;
        }
    }
}
