using System.Collections.ObjectModel;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public interface IDocumentTranslationCoordinator
{
    PauseController PauseController { get; }
    Task RunAsync(ObservableCollection<DocumentTranslationItem> items, AppSettings settings, CancellationToken cancellationToken);
}
