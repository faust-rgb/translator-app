using System.IO;
using System.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Services;
using TranslatorApp.Services.Ai;
using TranslatorApp.Services.Documents;
using TranslatorApp.ViewModels;

namespace TranslatorApp;

public partial class App : Application
{
    private IHost? _host;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        PdfSharpFontResolver.Initialize();

        var appDataDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TranslatorApp");
        Directory.CreateDirectory(appDataDirectory);

        _host = Host.CreateDefaultBuilder()
            .ConfigureAppConfiguration(builder =>
            {
                builder.Sources.Clear();
                builder.SetBasePath(AppContext.BaseDirectory);
                builder.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
                builder.AddJsonFile(Path.Combine(appDataDirectory, "user-settings.json"), optional: true, reloadOnChange: true);
            })
            .ConfigureServices((context, services) =>
            {
                services.Configure<AppSettings>(context.Configuration);

                services.AddHttpClient();

                services.AddSingleton<ISettingsService, SettingsService>();
                services.AddSingleton<IAppLogService, AppLogService>();
                services.AddSingleton<IConnectionTestService, ConnectionTestService>();
                services.AddSingleton<IGlossaryService, GlossaryService>();
                services.AddSingleton<IBilingualExportService, BilingualExportService>();
                services.AddSingleton<ITranslationHistoryService, TranslationHistoryService>();
                services.AddSingleton<IRecoveryStateService, RecoveryStateService>();
                services.AddSingleton<ITranslationProgressService, TranslationProgressService>();
                services.AddSingleton<IOcrService, OcrService>();
                services.AddSingleton<IAiTranslationClientFactory, AiTranslationClientFactory>();
                services.AddSingleton<ITranslationPromptBuilder, TranslationPromptBuilder>();
                services.AddSingleton<ITextTranslationService, TextTranslationService>();
                services.AddSingleton<IDocumentTranslator, WordDocumentTranslator>();
                services.AddSingleton<IDocumentTranslator, ExcelDocumentTranslator>();
                services.AddSingleton<IDocumentTranslator, PowerPointDocumentTranslator>();
                services.AddSingleton<IDocumentTranslator, PdfDocumentTranslator>();
                services.AddSingleton<IDocumentTranslationCoordinator, DocumentTranslationCoordinator>();
                services.AddSingleton<MainViewModel>();
                services.AddSingleton<MainWindow>();
            })
            .Build();

        await _host.StartAsync();

        var mainWindow = _host.Services.GetRequiredService<MainWindow>();
        mainWindow.DataContext = _host.Services.GetRequiredService<MainViewModel>();
        mainWindow.Show();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        if (_host is not null)
        {
            await _host.StopAsync();
            _host.Dispose();
        }

        base.OnExit(e);
    }
}
