namespace TranslatorApp.Infrastructure;

public static class TextDistributionHelper
{
    public static IReadOnlyList<string> Distribute(string translatedText, IReadOnlyList<int> weights)
    {
        if (weights.Count == 0)
        {
            return Array.Empty<string>();
        }

        if (weights.Count == 1)
        {
            return new[] { translatedText };
        }

        var totalWeight = Math.Max(1, weights.Sum(x => Math.Max(1, x)));
        var values = new string[weights.Count];
        var cursor = 0;

        for (var i = 0; i < weights.Count; i++)
        {
            if (i == weights.Count - 1)
            {
                values[i] = translatedText[cursor..];
                break;
            }

            var segmentLength = (int)Math.Round(translatedText.Length * (Math.Max(1, weights[i]) / (double)totalWeight));
            var nextCursor = Math.Clamp(cursor + segmentLength, cursor, translatedText.Length);
            values[i] = translatedText[cursor..nextCursor];
            cursor = nextCursor;
        }

        return values;
    }
}
