using System.Collections.ObjectModel;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;
using TranslatorApp.Services;
using TranslatorApp.Services.Ai;
using TranslatorApp.Services.Documents;

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: TranslatorCliRunner <document-path> [output-directory]");
    return 1;
}

var sourcePath = Path.GetFullPath(args[0]);
if (!File.Exists(sourcePath))
{
    Console.Error.WriteLine($"Source file not found: {sourcePath}");
    return 2;
}

var outputDirectory = args.Length > 1
    ? Path.GetFullPath(args[1])
    : string.Empty;

PdfSharpFontResolver.Initialize();

var appDataDirectory = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    "TranslatorApp");
Directory.CreateDirectory(appDataDirectory);

using var host = Host.CreateDefaultBuilder()
    .ConfigureAppConfiguration(builder =>
    {
        builder.Sources.Clear();
        builder.SetBasePath(AppContext.BaseDirectory);
        builder.AddJsonFile("appsettings.json", optional: true, reloadOnChange: false);
        builder.AddJsonFile(Path.Combine(appDataDirectory, "user-settings.json"), optional: true, reloadOnChange: false);
    })
    .ConfigureServices((context, services) =>
    {
        services.Configure<AppSettings>(context.Configuration);

        services.AddHttpClient();

        services.AddSingleton<ISecureApiKeyStorage, SecureApiKeyStorage>();
        services.AddSingleton<ISettingsService, SettingsService>();
        services.AddSingleton<IAppLogService, ConsoleAppLogService>();
        services.AddSingleton<IConnectionTestService, ConnectionTestService>();
        services.AddSingleton<IGlossaryService, GlossaryService>();
        services.AddSingleton<IBilingualExportService, BilingualExportService>();
        services.AddSingleton<ITranslationProgressService, ConsoleTranslationProgressService>();
        services.AddSingleton<ITranslationRequestThrottle, TranslationRequestThrottle>();
        services.AddSingleton<IOcrService, OcrService>();
        services.AddSingleton<IAiTranslationClientFactory, AiTranslationClientFactory>();
        services.AddSingleton<ITranslationPromptBuilder, TranslationPromptBuilder>();
        services.AddSingleton<ITextTranslationService, TextTranslationService>();
        services.AddSingleton<IDocumentTranslator, PdfDocumentTranslator>();
    })
    .Build();

await host.StartAsync();

var logService = host.Services.GetRequiredService<IAppLogService>();
logService.LogAdded += (_, message) => Console.WriteLine(message);

var settingsService = host.Services.GetRequiredService<ISettingsService>();
var settings = await settingsService.LoadAsync();
if (!string.IsNullOrWhiteSpace(outputDirectory))
{
    settings.Translation.OutputDirectory = outputDirectory;
}

if (string.IsNullOrWhiteSpace(settings.Ai.ApiKey))
{
    Console.Error.WriteLine("No API key found in saved settings / secure storage.");
    return 3;
}

var item = new DocumentTranslationItem
{
    SourcePath = sourcePath,
    FileType = Path.GetExtension(sourcePath).TrimStart('.').ToUpperInvariant(),
    Status = DocumentStatus.Pending,
    Progress = 0,
    ProgressText = "Queued from CLI"
};

using var cts = new CancellationTokenSource();
var translator = host.Services
    .GetRequiredService<IEnumerable<IDocumentTranslator>>()
    .FirstOrDefault(x => x.CanHandle(Path.GetExtension(sourcePath).ToLowerInvariant()));

if (translator is null)
{
    Console.Error.WriteLine($"No translator found for: {sourcePath}");
    return 5;
}

item.Status = DocumentStatus.Running;
item.ProgressText = "Starting";

var context = new TranslationJobContext
{
    Settings = settings,
    Item = item,
    CancellationToken = cts.Token,
    PauseController = new PauseController(),
    ResumeUnitIndex = 0,
    ResumeSubUnitIndex = 0,
    ReportProgressAsync = (progress, text) =>
    {
        item.Progress = progress;
        item.ProgressText = text;
        Console.WriteLine($"Progress: {progress}% {text}");
        return Task.CompletedTask;
    },
    SaveCheckpointAsync = (_, _, _) => Task.CompletedTask
};

try
{
    await translator.TranslateAsync(context);
    item.Status = DocumentStatus.Completed;
    item.Progress = 100;
    item.ProgressText = "Completed";
}
catch (Exception ex)
{
    item.Status = DocumentStatus.Failed;
    item.ErrorMessage = ex.Message;
}

Console.WriteLine($"Final status: {item.Status}");
Console.WriteLine($"Output path: {item.OutputPath}");
if (!string.IsNullOrWhiteSpace(item.ErrorMessage))
{
    Console.Error.WriteLine(item.ErrorMessage);
}

await host.StopAsync();
return item.Status == DocumentStatus.Completed ? 0 : 4;

file sealed class ConsoleAppLogService : IAppLogService
{
    public event EventHandler<string>? LogAdded;

    public void Info(string message) => Publish($"INFO  {message}");

    public void Error(string message) => Publish($"ERROR {message}");

    private void Publish(string message)
    {
        Console.WriteLine(message);
        LogAdded?.Invoke(this, message);
    }
}

file sealed class ConsoleTranslationProgressService : ITranslationProgressService
{
    public event EventHandler<TranslationProgressEventArgs>? StreamUpdated;

    public void Publish(string title, string partialText)
    {
        StreamUpdated?.Invoke(this, new TranslationProgressEventArgs(title, partialText));
    }

    public void Clear() => Publish(string.Empty, string.Empty);
}
