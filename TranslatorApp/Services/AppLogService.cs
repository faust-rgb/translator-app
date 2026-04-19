using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;

namespace TranslatorApp.Services;

public sealed class AppLogService : IAppLogService
{
    private readonly object _syncRoot = new();
    private readonly string _logFilePath;

    public AppLogService(string appDataDirectory)
    {
        var logDirectory = Path.Combine(appDataDirectory, "Logs");
        Directory.CreateDirectory(logDirectory);
        _logFilePath = Path.Combine(logDirectory, $"TranslatorApp-{DateTime.Now:yyyyMMdd}.log");
    }

    public event EventHandler<string>? LogAdded;

    public void Info(string message) => Publish($"[{FormatTimestamp()}] INFO  {message}");

    public void Error(string message) => Publish($"[{FormatTimestamp()}] ERROR {message}");

    private void Publish(string message)
    {
        WriteToFile(message);

        if (Application.Current?.Dispatcher.CheckAccess() == true)
        {
            LogAdded?.Invoke(this, message);
            return;
        }

        if (Application.Current?.Dispatcher is { } dispatcher)
        {
            dispatcher.Invoke(() => LogAdded?.Invoke(this, message));
            return;
        }

        LogAdded?.Invoke(this, message);
    }

    private static string FormatTimestamp() => DateTimeOffset.UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'");

    private void WriteToFile(string message)
    {
        try
        {
            lock (_syncRoot)
            {
                File.AppendAllText(_logFilePath, message + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to write log file '{_logFilePath}': {ex}");
        }
    }
}
