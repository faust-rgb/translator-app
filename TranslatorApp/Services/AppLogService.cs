using System.Windows;

namespace TranslatorApp.Services;

public sealed class AppLogService : IAppLogService
{
    public event EventHandler<string>? LogAdded;

    public void Info(string message) => Publish($"[{FormatTimestamp()}] INFO  {message}");

    public void Error(string message) => Publish($"[{FormatTimestamp()}] ERROR {message}");

    private void Publish(string message)
    {
        if (Application.Current?.Dispatcher.CheckAccess() == true)
        {
            LogAdded?.Invoke(this, message);
            return;
        }

        Application.Current?.Dispatcher.Invoke(() => LogAdded?.Invoke(this, message));
    }

    private static string FormatTimestamp() => DateTimeOffset.UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'");
}
