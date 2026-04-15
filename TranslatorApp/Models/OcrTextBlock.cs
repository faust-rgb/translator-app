namespace TranslatorApp.Models;

public sealed record OcrTextBlock(string Text, double Left, double Top, double Width, double Height);
