using System.Diagnostics;
using System.IO;
using System.Text;

namespace TranslatorApp.Services.Documents;

public sealed class CalibreEbookConversionService(IAppLogService logService) : IEbookConversionService
{
    private static readonly string[] KnownExecutableLocations =
    [
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Calibre2", "ebook-convert.exe"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Calibre2", "ebook-convert.exe")
    ];

    public async Task ConvertAsync(string inputPath, string outputPath, string configuredExecutablePath, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var executablePath = ResolveExecutablePath(configuredExecutablePath);
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        logService.Info($"正在调用 Calibre 转换：{Path.GetFileName(inputPath)} -> {Path.GetFileName(outputPath)}");

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = executablePath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        process.StartInfo.ArgumentList.Add(inputPath);
        process.StartInfo.ArgumentList.Add(outputPath);

        var outputBuilder = new StringBuilder();
        var errorBuilder = new StringBuilder();

        process.OutputDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                outputBuilder.AppendLine(args.Data);
            }
        };
        process.ErrorDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                errorBuilder.AppendLine(args.Data);
            }
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        using var registration = cancellationToken.Register(() =>
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // ignore cancellation cleanup failures
            }
        });

        await process.WaitForExitAsync(cancellationToken);

        if (process.ExitCode != 0 || !File.Exists(outputPath))
        {
            var details = string.Join(
                Environment.NewLine,
                new[] { errorBuilder.ToString().Trim(), outputBuilder.ToString().Trim() }
                    .Where(static x => !string.IsNullOrWhiteSpace(x)));

            throw new InvalidOperationException(
                "Calibre 转换失败。请确认已安装 Calibre，并且 `ebook-convert.exe` 路径有效。" +
                (string.IsNullOrWhiteSpace(details) ? string.Empty : $"{Environment.NewLine}{details}"));
        }
    }

    private static string ResolveExecutablePath(string configuredExecutablePath)
    {
        if (!string.IsNullOrWhiteSpace(configuredExecutablePath))
        {
            if (File.Exists(configuredExecutablePath))
            {
                return configuredExecutablePath;
            }

            throw new FileNotFoundException("配置的 ebook-convert.exe 不存在。", configuredExecutablePath);
        }

        var envPath = Environment.GetEnvironmentVariable("CALIBRE_EBOOK_CONVERT");
        if (!string.IsNullOrWhiteSpace(envPath) && File.Exists(envPath))
        {
            return envPath;
        }

        foreach (var location in KnownExecutableLocations.Where(File.Exists))
        {
            return location;
        }

        var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var candidate = Path.Combine(directory, "ebook-convert.exe");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new FileNotFoundException("未找到 Calibre 的 ebook-convert.exe。请先安装 Calibre，或在“翻译设置”中手动指定路径。");
    }
}
