using System.IO;
using System.Threading;
using TranslatorApp.Configuration;
using TranslatorApp.Infrastructure;
using TranslatorApp.Models;

namespace TranslatorApp.Services.Documents;

public abstract class DocumentTranslatorBase(ITextTranslationService textTranslationService, IAppLogService logService)
    : IDocumentTranslator
{
    public abstract bool CanHandle(string extension);

    public abstract Task TranslateAsync(TranslationJobContext context);

    protected async Task<string> TranslateBlockAsync(
        string text,
        string contextHint,
        string additionalRequirements,
        AppSettings settings,
        PauseController pauseController,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        await pauseController.WaitIfPausedAsync(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return await textTranslationService.TranslateAsync(text, contextHint, additionalRequirements, settings, onPartialResponse, cancellationToken);
    }

    protected async Task<IReadOnlyList<string>> TranslateBatchAsync(
        IReadOnlyList<TranslationBlock> blocks,
        AppSettings settings,
        PauseController pauseController,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        if (blocks.Count == 0)
        {
            return Array.Empty<string>();
        }

        if (TryUsePackedBatchMode(blocks))
        {
            return await TranslatePackedBatchAsync(blocks, settings, pauseController, onPartialResponse, cancellationToken);
        }

        var results = new string[blocks.Count];
        var maxConcurrency = Math.Min(blocks.Count, GetBlockTranslationConcurrency(settings));
        if (maxConcurrency <= 1)
        {
            for (var i = 0; i < blocks.Count; i++)
            {
                results[i] = await TranslateBlockAsync(
                    blocks[i].Text,
                    blocks[i].ContextHint,
                    blocks[i].AdditionalRequirements,
                    settings,
                    pauseController,
                    onPartialResponse,
                    cancellationToken);
            }

            return results;
        }

        using var semaphore = new SemaphoreSlim(maxConcurrency);
        var previewAssigned = 0;

        var tasks = blocks.Select(async (block, index) =>
        {
            await semaphore.WaitAsync(cancellationToken);
            try
            {
                Func<string, Task>? partialCallback = null;
                if (onPartialResponse is not null &&
                    Interlocked.CompareExchange(ref previewAssigned, 1, 0) == 0)
                {
                    partialCallback = onPartialResponse;
                }

                results[index] = await TranslateBlockAsync(
                    block.Text,
                    block.ContextHint,
                    block.AdditionalRequirements,
                    settings,
                    pauseController,
                    partialCallback,
                    cancellationToken);
            }
            finally
            {
                semaphore.Release();
            }
        });

        await Task.WhenAll(tasks);
        return results;
    }

    private async Task<IReadOnlyList<string>> TranslatePackedBatchAsync(
        IReadOnlyList<TranslationBlock> blocks,
        AppSettings settings,
        PauseController pauseController,
        Func<string, Task>? onPartialResponse,
        CancellationToken cancellationToken)
    {
        var packedGroups = BuildPackedGroups(blocks);
        var results = new string[blocks.Count];
        var previewAssigned = false;

        foreach (var group in packedGroups)
        {
            if (group.Count == 1)
            {
                var single = group.Items[0];
                Func<string, Task>? partialCallback = null;
                if (!previewAssigned && onPartialResponse is not null)
                {
                    previewAssigned = true;
                    partialCallback = onPartialResponse;
                }

                results[single.Index] = await TranslateBlockAsync(
                    single.Block.Text,
                    single.Block.ContextHint,
                    single.Block.AdditionalRequirements,
                    settings,
                    pauseController,
                    partialCallback,
                    cancellationToken);
                continue;
            }

            var packedRequest = BuildPackedRequest(group.Items);
            Func<string, Task>? packedPartialCallback = null;
            if (!previewAssigned && onPartialResponse is not null)
            {
                previewAssigned = true;
                packedPartialCallback = onPartialResponse;
            }

            var packedTranslation = await TranslateBlockAsync(
                packedRequest.Text,
                packedRequest.ContextHint,
                packedRequest.AdditionalRequirements,
                settings,
                pauseController,
                packedPartialCallback,
                cancellationToken);

            if (!TryParsePackedTranslation(group.Items, packedTranslation, out var parsed))
            {
                foreach (var item in group.Items)
                {
                    results[item.Index] = await TranslateBlockAsync(
                        item.Block.Text,
                        item.Block.ContextHint,
                        item.Block.AdditionalRequirements,
                        settings,
                        pauseController,
                        null,
                        cancellationToken);
                }

                continue;
            }

            foreach (var item in group.Items)
            {
                results[item.Index] = parsed[item.Index];
            }
        }

        return results;
    }

    private static bool TryUsePackedBatchMode(IReadOnlyList<TranslationBlock> blocks)
    {
        if (blocks.Count < 2)
        {
            return false;
        }

        return blocks.Any(block => IsPackableBlock(block)) &&
               blocks.All(block => !string.IsNullOrWhiteSpace(block.Text));
    }

    private static List<PackedGroup> BuildPackedGroups(IReadOnlyList<TranslationBlock> blocks)
    {
        var groups = new List<PackedGroup>();
        var current = new List<IndexedTranslationBlock>();
        var currentRequirementSignature = string.Empty;
        var currentProfile = PackingProfile.Default;
        var currentTotalLength = 0;

        for (var index = 0; index < blocks.Count; index++)
        {
            var item = new IndexedTranslationBlock(index, blocks[index]);
            if (!IsPackableBlock(item.Block))
            {
                FlushCurrent();
                groups.Add(new PackedGroup([item]));
                continue;
            }

            var profile = ResolvePackingProfile(item.Block);
            var requirementSignature = BuildPackingRequirementSignature(item.Block, profile);
            var blockLength = item.Block.Text.Length;
            var sameRequirementGroup = current.Count == 0 ||
                                       (string.Equals(currentProfile.Name, profile.Name, StringComparison.Ordinal) &&
                                        string.Equals(currentRequirementSignature, requirementSignature, StringComparison.Ordinal));
            var fitsGroup = current.Count < profile.MaxItems && currentTotalLength + blockLength <= profile.MaxChars;

            if (!sameRequirementGroup || !fitsGroup)
            {
                FlushCurrent();
            }

            if (current.Count == 0)
            {
                currentProfile = profile;
                currentRequirementSignature = requirementSignature;
            }

            current.Add(item);
            currentTotalLength += blockLength;
        }

        FlushCurrent();
        return groups;

        void FlushCurrent()
        {
            if (current.Count == 0)
            {
                return;
            }

            groups.Add(new PackedGroup(current.ToList()));
            current.Clear();
            currentRequirementSignature = string.Empty;
            currentProfile = PackingProfile.Default;
            currentTotalLength = 0;
        }
    }

    private static bool IsPackableBlock(TranslationBlock block)
    {
        var text = block.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text) || text.Length > 900)
        {
            return false;
        }

        if (text.Contains("[[BATCH_BLOCK_", StringComparison.Ordinal))
        {
            return false;
        }

        var requirements = block.AdditionalRequirements ?? string.Empty;
        return !requirements.Contains("严格保留类似 [[SEG_", StringComparison.Ordinal) &&
               !requirements.Contains("类型：代码", StringComparison.Ordinal) &&
               !requirements.Contains("类型：预格式化代码", StringComparison.Ordinal);
    }

    private static PackingProfile ResolvePackingProfile(TranslationBlock block)
    {
        var context = block.ContextHint ?? string.Empty;
        var requirements = block.AdditionalRequirements ?? string.Empty;
        var textLength = (block.Text ?? string.Empty).Length;

        if (context.Contains("EPUB 目录", StringComparison.Ordinal))
        {
            return new PackingProfile("epub-nav", 12, 3400);
        }

        if (context.Contains("Excel", StringComparison.Ordinal) && textLength <= 220)
        {
            return new PackingProfile("excel-short", 10, 3200);
        }

        if (context.Contains("PowerPoint", StringComparison.Ordinal) && textLength <= 220)
        {
            return new PackingProfile("ppt-short", 8, 3000);
        }

        if ((context.Contains("Word 表格单元格", StringComparison.Ordinal) ||
             requirements.Contains("类型：表格单元格", StringComparison.Ordinal)) &&
            textLength <= 240)
        {
            return new PackingProfile("word-table-cell", 8, 2800);
        }

        if ((context.Contains("Word 列表项", StringComparison.Ordinal) ||
             requirements.Contains("类型：列表", StringComparison.Ordinal)) &&
            textLength <= 220)
        {
            return new PackingProfile("word-list", 8, 2600);
        }

        if ((context.Contains("Word 标题段落", StringComparison.Ordinal) ||
             requirements.Contains("类型：标题", StringComparison.Ordinal)) &&
            textLength <= 180)
        {
            return new PackingProfile("word-heading", 8, 2200);
        }

        return PackingProfile.Default;
    }

    private static string BuildPackingRequirementSignature(TranslationBlock block, PackingProfile profile)
    {
        var requirements = (block.AdditionalRequirements ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(requirements))
        {
            return profile.Name;
        }

        return profile.Name switch
        {
            "epub-nav" => "epub-nav",
            "excel-short" => profile.Name + "::" + ExtractLeadingRequirementLine(requirements),
            "ppt-short" => profile.Name + "::" + ExtractLeadingRequirementLine(requirements),
            "word-table-cell" => profile.Name + "::" + ExtractLeadingRequirementLine(requirements),
            "word-list" => profile.Name + "::" + ExtractLeadingRequirementLine(requirements),
            "word-heading" => profile.Name + "::" + ExtractLeadingRequirementLine(requirements),
            _ => profile.Name + "::" + requirements
        };
    }

    private static string ExtractLeadingRequirementLine(string requirements)
    {
        var firstLine = requirements
            .Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault();
        return firstLine ?? string.Empty;
    }

    private static PackedTranslationRequest BuildPackedRequest(IReadOnlyList<IndexedTranslationBlock> items)
    {
        var builder = new System.Text.StringBuilder();
        var contexts = new List<string>();

        foreach (var item in items)
        {
            var marker = BuildBatchBlockMarker(item.Index);
            builder.Append(marker);
            builder.AppendLine();
            builder.Append(item.Block.Text.Trim());
            builder.AppendLine();
            contexts.Add($"{marker} => {item.Block.ContextHint}");
        }

        var contextHint =
            "以下是同一文档中的多个独立片段，请分别翻译并保留各自标记；不要互相合并、重排或遗漏任何片段。" +
            Environment.NewLine +
            string.Join(Environment.NewLine, contexts);

        var additionalRequirements =
            string.IsNullOrWhiteSpace(items[0].Block.AdditionalRequirements)
                ? string.Empty
                : items[0].Block.AdditionalRequirements.Trim() + Environment.NewLine;

        additionalRequirements +=
            "当前请求包含多个独立片段。请严格保留并原样输出每个 [[BATCH_BLOCK_xxx]] 标记，" +
            "每个标记后只跟对应片段的译文；不要增删标记，不要交换顺序，也不要把多个片段合并成一个片段。";

        return new PackedTranslationRequest(builder.ToString().TrimEnd(), contextHint, additionalRequirements);
    }

    private static bool TryParsePackedTranslation(
        IReadOnlyList<IndexedTranslationBlock> items,
        string packedTranslation,
        out Dictionary<int, string> parsed)
    {
        parsed = new Dictionary<int, string>();
        if (string.IsNullOrWhiteSpace(packedTranslation))
        {
            return false;
        }

        for (var i = 0; i < items.Count; i++)
        {
            var marker = BuildBatchBlockMarker(items[i].Index);
            var start = packedTranslation.IndexOf(marker, StringComparison.Ordinal);
            if (start < 0)
            {
                parsed.Clear();
                return false;
            }

            start += marker.Length;
            while (start < packedTranslation.Length &&
                   (packedTranslation[start] == '\r' || packedTranslation[start] == '\n' || char.IsWhiteSpace(packedTranslation[start])))
            {
                start++;
            }

            var end = packedTranslation.Length;
            for (var next = i + 1; next < items.Count; next++)
            {
                var nextMarker = BuildBatchBlockMarker(items[next].Index);
                var nextPosition = packedTranslation.IndexOf(nextMarker, start, StringComparison.Ordinal);
                if (nextPosition >= 0)
                {
                    end = nextPosition;
                    break;
                }
            }

            parsed[items[i].Index] = packedTranslation[start..end].Trim();
        }

        return parsed.Count == items.Count;
    }

    private static string BuildBatchBlockMarker(int index) => $"[[BATCH_BLOCK_{index + 1:000}]]";

    protected static string BuildOutputPath(string sourcePath, string outputDirectory, string? outputExtension = null)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var extension = string.IsNullOrWhiteSpace(outputExtension)
            ? Path.GetExtension(sourcePath)
            : outputExtension;
        var directory = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.GetDirectoryName(sourcePath)!
            : outputDirectory;

        Directory.CreateDirectory(directory);
        return Path.Combine(directory, $"{fileName}.translated{extension}");
    }

    protected void Log(string message) => logService.Info(message);

    protected static int GetBlockTranslationConcurrency(AppSettings settings) =>
        Math.Max(1, settings.Translation.MaxParallelBlocks);

    protected static (int Start, int End) GetRequestedRange(AppSettings settings, int totalUnits)
    {
        totalUnits = Math.Max(0, totalUnits);
        if (totalUnits == 0)
        {
            return (1, 0);
        }

        var start = Math.Max(1, settings.Translation.RangeStart);
        var end = settings.Translation.RangeEnd <= 0 ? totalUnits : Math.Max(start, settings.Translation.RangeEnd);
        start = Math.Min(start, totalUnits);
        end = Math.Min(end, totalUnits);
        return (start, end);
    }

    protected static bool IsWithinRequestedRange(int oneBasedIndex, (int Start, int End) range) =>
        range.End >= range.Start && oneBasedIndex >= range.Start && oneBasedIndex <= range.End;

    protected static string DescribeRequestedRange(string unitName, (int Start, int End) range) =>
        range.End < range.Start ? $"未命中任何{unitName}" : $"{unitName} {range.Start}-{range.End}";

    protected readonly record struct TranslationBlock(string Text, string ContextHint, string AdditionalRequirements = "");

    private readonly record struct IndexedTranslationBlock(int Index, TranslationBlock Block);

    private readonly record struct PackedGroup(IReadOnlyList<IndexedTranslationBlock> Items)
    {
        public int Count => Items.Count;
    }

    private readonly record struct PackedTranslationRequest(string Text, string ContextHint, string AdditionalRequirements);

    private readonly record struct PackingProfile(string Name, int MaxItems, int MaxChars)
    {
        public static PackingProfile Default { get; } = new("default", 6, 2600);
    }
}
