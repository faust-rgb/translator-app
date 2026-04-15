using System.Collections.Concurrent;
using System.IO;
using TranslatorApp.Models;

namespace TranslatorApp.Services;

public sealed class GlossaryService(IAppLogService logService) : IGlossaryService, IDisposable
{
    private readonly ConcurrentDictionary<string, IReadOnlyList<GlossaryEntry>> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, FileSystemWatcher> _watchers = new(StringComparer.OrdinalIgnoreCase);

    public async Task<IReadOnlyList<GlossaryEntry>> LoadAsync(string glossaryPath, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(glossaryPath) || !File.Exists(glossaryPath))
        {
            return Array.Empty<GlossaryEntry>();
        }

        EnsureWatcher(glossaryPath);

        if (_cache.TryGetValue(glossaryPath, out var cached))
        {
            return cached;
        }

        var loaded = await LoadEntriesAsync(glossaryPath, cancellationToken);
        _cache[glossaryPath] = loaded;
        return loaded;
    }

    public string BuildPromptSection(string text, IReadOnlyList<GlossaryEntry> entries)
    {
        var matched = entries
            .Where(x => text.Contains(x.Source, StringComparison.OrdinalIgnoreCase))
            .Take(30)
            .ToList();

        if (matched.Count == 0)
        {
            return string.Empty;
        }

        return "术语表约束：\n" + string.Join('\n', matched.Select(x => $"- {x.Source} => {x.Target}"));
    }

    public void Dispose()
    {
        foreach (var watcher in _watchers.Values)
        {
            watcher.Dispose();
        }
    }

    private void EnsureWatcher(string glossaryPath)
    {
        _watchers.GetOrAdd(glossaryPath, path =>
        {
            var watcher = new FileSystemWatcher(Path.GetDirectoryName(path)!, Path.GetFileName(path))
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName
            };

            void Reload(object? _, FileSystemEventArgs __)
            {
                try
                {
                    if (!File.Exists(path))
                    {
                        _cache[path] = Array.Empty<GlossaryEntry>();
                        return;
                    }

                    _cache[path] = LoadEntriesAsync(path, CancellationToken.None).GetAwaiter().GetResult();
                    logService.Info($"术语表已热加载：{Path.GetFileName(path)}");
                }
                catch (Exception ex)
                {
                    logService.Error($"术语表重载失败：{ex.Message}");
                }
            }

            watcher.Changed += Reload;
            watcher.Created += Reload;
            watcher.Renamed += (_, __) => Reload(_, null!);
            watcher.EnableRaisingEvents = true;
            return watcher;
        });
    }

    private static async Task<IReadOnlyList<GlossaryEntry>> LoadEntriesAsync(string glossaryPath, CancellationToken cancellationToken)
    {
        var lines = await File.ReadAllLinesAsync(glossaryPath, cancellationToken);
        return lines
            .Select(ParseLine)
            .OfType<GlossaryEntry>()
            .ToList();
    }

    private static GlossaryEntry? ParseLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith('#'))
        {
            return null;
        }

        var parts = line.Contains('\t')
            ? line.Split('\t', 2, StringSplitOptions.TrimEntries)
            : line.Split('=', 2, StringSplitOptions.TrimEntries);

        return parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[0]) && !string.IsNullOrWhiteSpace(parts[1])
            ? new GlossaryEntry(parts[0], parts[1])
            : null;
    }
}
