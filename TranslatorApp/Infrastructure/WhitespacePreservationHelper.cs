namespace TranslatorApp.Infrastructure;

public static class WhitespacePreservationHelper
{
    public static string GetLeadingWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var index = 0;
        while (index < text.Length && char.IsWhiteSpace(text[index]))
        {
            index++;
        }

        return text[..index];
    }

    public static string GetTrailingWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var index = text.Length - 1;
        while (index >= 0 && char.IsWhiteSpace(text[index]))
        {
            index--;
        }

        return text[(index + 1)..];
    }

    public static string PreserveEdgeWhitespace(string original, string translated)
    {
        ArgumentNullException.ThrowIfNull(original);
        translated ??= string.Empty;

        var leading = GetLeadingWhitespace(original);
        var trailing = GetTrailingWhitespace(original);
        return string.Concat(leading, translated.Trim(), trailing);
    }
}
