using TranslatorApp.Configuration;

namespace TranslatorApp.Services;

public interface IConnectionTestService
{
    Task TestAsync(AppSettings settings, CancellationToken cancellationToken);
}
