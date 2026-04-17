using System.Text;

namespace TranslatorApp.Infrastructure;

public static class TextDistributionHelper
{
    /// <summary>
    /// 将翻译后的文本智能分配到多个目标段落。
    /// 使用语义边界感知算法，优先在句子、分句、词语边界处分割。
    /// </summary>
    public static IReadOnlyList<string> Distribute(string translatedText, IReadOnlyList<int> weights)
    {
        ArgumentNullException.ThrowIfNull(weights);
        translatedText ??= string.Empty;

        if (weights.Count == 0)
        {
            return Array.Empty<string>();
        }

        if (weights.Any(w => w < 0))
        {
            throw new ArgumentException("权重不能为负数。", nameof(weights));
        }

        if (weights.Count == 1)
        {
            return new[] { translatedText };
        }

        // 如果译文为空，返回空字符串列表
        if (string.IsNullOrEmpty(translatedText))
        {
            return weights.Select(_ => string.Empty).ToList();
        }

        // 计算目标长度比例
        var totalWeight = Math.Max(1, weights.Sum(x => Math.Max(1, x)));
        var targetLengths = weights.Select(w => translatedText.Length * (Math.Max(1, w) / (double)totalWeight)).ToList();

        // 使用语义边界感知算法分割文本
        var segments = SplitBySemanticBoundaries(translatedText, targetLengths);

        return segments;
    }

    /// <summary>
    /// 基于语义边界分割文本。优先顺序：
    /// 1. 句子边界（句号、问号、感叹号等）
    /// 2. 分句边界（逗号、分号等）
    /// 3. 词语边界（空格、标点等）
    /// 4. 字符边界（不得已时）
    /// </summary>
    private static IReadOnlyList<string> SplitBySemanticBoundaries(string text, IReadOnlyList<double> targetLengths)
    {
        var segments = new List<string>();
        var remainingText = text;
        var startIndex = 0;

        for (var i = 0; i < targetLengths.Count - 1; i++)
        {
            var targetEnd = startIndex + targetLengths[i];
            var splitPoint = FindBestSplitPoint(text, startIndex, targetEnd, targetLengths[i] * 0.3);

            segments.Add(text[startIndex..splitPoint]);
            startIndex = splitPoint;
        }

        // 最后一段取剩余所有内容
        segments.Add(text[startIndex..]);

        // 如果段数不足，补充空字符串
        while (segments.Count < targetLengths.Count)
        {
            segments.Add(string.Empty);
        }

        return segments;
    }

    /// <summary>
    /// 在指定范围内寻找最佳分割点。
    /// </summary>
    /// <param name="text">完整文本</param>
    /// <param name="startIndex">起始索引</param>
    /// <param name="targetEnd">目标结束位置</param>
    /// <param name="tolerance">容差范围（允许前后调整的幅度）</param>
    /// <returns>最佳分割点的索引</returns>
    private static int FindBestSplitPoint(string text, int startIndex, double targetEnd, double tolerance)
    {
        var minIndex = (int)Math.Max(startIndex, Math.Floor(targetEnd - tolerance));
        var maxIndex = (int)Math.Min(text.Length, Math.Ceiling(targetEnd + tolerance));

        // 确保范围有效
        minIndex = Math.Max(startIndex, minIndex);
        maxIndex = Math.Min(text.Length, maxIndex);

        if (minIndex >= maxIndex)
        {
            return Math.Min((int)targetEnd, text.Length);
        }

        // 在范围内搜索最佳分割点
        var searchRange = text[minIndex..maxIndex];

        // 优先级1：句子结束标点（。！？.!?）
        var sentenceEnds = new[] { '。', '！', '？', '.', '!', '?' };
        for (var i = searchRange.Length - 1; i >= 0; i--)
        {
            if (sentenceEnds.Contains(searchRange[i]))
            {
                return minIndex + i + 1;
            }
        }

        // 优先级2：分句标点（，；,;）
        var clauseEnds = new[] { '，', '；', ',', ';', '、' };
        for (var i = searchRange.Length - 1; i >= 0; i--)
        {
            if (clauseEnds.Contains(searchRange[i]))
            {
                return minIndex + i + 1;
            }
        }

        // 优先级3：空格（适用于英文等空格分隔语言）
        for (var i = searchRange.Length - 1; i >= 0; i--)
        {
            if (char.IsWhiteSpace(searchRange[i]))
            {
                return minIndex + i + 1;
            }
        }

        // 优先级4：其他标点符号
        for (var i = searchRange.Length - 1; i >= 0; i--)
        {
            if (char.IsPunctuation(searchRange[i]))
            {
                return minIndex + i + 1;
            }
        }

        // 优先级5：字符边界（CJK字符）
        for (var i = searchRange.Length - 1; i >= 0; i--)
        {
            var ch = searchRange[i];
            // CJK统一表意文字范围
            if (IsCjkCharacter(ch))
            {
                return minIndex + i + 1;
            }
        }

        // 回退：直接按目标长度分割
        return Math.Min((int)targetEnd, text.Length);
    }

    /// <summary>
    /// 判断字符是否为CJK（中日韩）统一表意文字。
    /// </summary>
    private static bool IsCjkCharacter(char ch)
    {
        // CJK统一表意文字基本区和扩展区A
        return (ch >= '\u4E00' && ch <= '\u9FFF') ||  // CJK统一表意文字
               (ch >= '\u3400' && ch <= '\u4DBF') ||  // CJK统一表意文字扩展区A
               (ch >= '\uF900' && ch <= '\uFAFF');    // CJK兼容表意文字
    }

    /// <summary>
    /// 将翻译后的文本分配到多个目标段落，保持原有语义结构。
    /// 这是简化版本，适用于需要保持原有段落数量的情况。
    /// </summary>
    public static IReadOnlyList<string> DistributePreservingStructure(string translatedText, int segmentCount)
    {
        if (segmentCount <= 0)
        {
            return Array.Empty<string>();
        }

        if (segmentCount == 1 || string.IsNullOrEmpty(translatedText))
        {
            return Enumerable.Repeat(translatedText ?? string.Empty, segmentCount).ToList();
        }

        var targetLength = translatedText.Length / (double)segmentCount;
        var targetLengths = Enumerable.Range(0, segmentCount).Select(_ => targetLength).ToList();

        return SplitBySemanticBoundaries(translatedText, targetLengths);
    }
}
