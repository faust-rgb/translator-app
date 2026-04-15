namespace TranslatorApp.Services;

public interface IAppLogService
{
    event EventHandler<string>? LogAdded;
    void Info(string message);
    void Error(string message);
}
