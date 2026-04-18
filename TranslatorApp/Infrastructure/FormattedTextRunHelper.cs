namespace TranslatorApp.Infrastructure;

public static class FormattedTextRunHelper
{
    public static IReadOnlyList<FormattedTextGroup<TText>> GroupAdjacentRunsByFormat<TText>(
        IReadOnlyList<FormattedTextRun<TText>> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);

        var groups = new List<FormattedTextGroup<TText>>();
        foreach (var run in runs)
        {
            if (groups.Count > 0 &&
                string.Equals(groups[^1].FormatKey, run.FormatKey, StringComparison.Ordinal))
            {
                groups[^1].Texts.AddRange(run.Texts);
                groups[^1].Original += run.Original;
                continue;
            }

            groups.Add(new FormattedTextGroup<TText>(run.FormatKey, run.Texts.ToList(), run.Original));
        }

        return groups;
    }

    public static IReadOnlyList<string> DistributeAcrossGroups<TText>(
        string translated,
        IReadOnlyList<FormattedTextGroup<TText>> groups)
    {
        ArgumentNullException.ThrowIfNull(groups);

        if (groups.Count == 0)
        {
            return Array.Empty<string>();
        }

        if (groups.Count == 1)
        {
            return new[] { translated ?? string.Empty };
        }

        return TextDistributionHelper.Distribute(
            translated,
            groups.Select(x => Math.Max(1, x.Original.Length)).ToList());
    }
}

public sealed class FormattedTextGroup<TText>(string formatKey, List<TText> texts, string original)
{
    public string FormatKey { get; } = formatKey;
    public List<TText> Texts { get; } = texts;
    public string Original { get; set; } = original;
}

public sealed record FormattedTextRun<TText>(IReadOnlyList<TText> Texts, string Original, string FormatKey);
