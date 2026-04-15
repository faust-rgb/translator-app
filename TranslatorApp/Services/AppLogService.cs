using System.Windows;

namespace TranslatorApp.Services;

public sealed class AppLogService : IAppLogService
{
    public event EventHandler<string>? LogAdded;

    public void Info(string message) => Publish($"[{DateTime.Now:HH:mm:ss}] INFO  {message}");

    public void Error(string message) => Publish($"[{DateTime.Now:HH:mm:ss}] ERROR {message}");

    private void Publish(string message)
    {
        if (Application.Current?.Dispatcher.CheckAccess() == true)
        {
            LogAdded?.Invoke(this, message);
            return;
        }

        Application.Current?.Dispatcher.Invoke(() => LogAdded?.Invoke(this, message));
    }
}
