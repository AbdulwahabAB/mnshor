using System.Collections.Concurrent;
using System.Globalization;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace Mnshor.Engines;

public static class EmojiRenderer
{
    private static readonly ConcurrentDictionary<string, byte[]> EmojiPngCache = new(StringComparer.Ordinal);

    private static readonly Lazy<SKTypeface> EmojiTypeface = new(() =>
    {
        string fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "seguiemj.ttf");
        if (File.Exists(fontPath))
        {
            var tfFromFile = SKTypeface.FromFile(fontPath);
            if (tfFromFile != null)
                return tfFromFile;
        }

        var tf = SKTypeface.FromFamilyName("Segoe UI Emoji");
        return tf ?? SKTypeface.Default;
    });

    /// <summary>
    /// Determines whether a Unicode grapheme cluster (text element) is an emoji or symbol that requires emoji rendering.
    /// </summary>
    public static bool IsEmojiCluster(string cluster)
    {
        if (string.IsNullOrEmpty(cluster))
            return false;

        int i = 0;
        while (i < cluster.Length)
        {
            int cp = char.ConvertToUtf32(cluster, i);
            i += char.IsSurrogatePair(cluster, i) ? 2 : 1;

            if (IsEmojiCodePoint(cp))
                return true;
        }

        return false;
    }

    public static bool IsEmojiCodePoint(int cp)
    {
        return (cp >= 0x1F000 && cp <= 0x1FAFF) // All SMP Emojis, Emoticons, Transport, Supplemental Symbols, Pictographs, Flags
            || (cp >= 0x2600 && cp <= 0x27BF)   // Misc Symbols & Dingbats (☀️, ✅, ❤️, ✨, etc.)
            || (cp >= 0x2300 && cp <= 0x23FF)   // Misc Technical (⌚, ⌛, ⏰, ⏳, etc.)
            || (cp >= 0x2B00 && cp <= 0x2BFF)   // Stars & Arrows (⭐, ⭕, ⬆️, etc.)
            || (cp >= 0x25A0 && cp <= 0x25FF)   // Geometric Shapes (▶️, ◀️, etc.)
            || (cp >= 0x2190 && cp <= 0x21FF)   // Arrows
            || cp == 0x200D                     // Zero Width Joiner
            || cp == 0x20E3                     // Combining Enclosing Keycap (1️⃣, #️⃣)
            || cp == 0xFE0F                     // Emoji Variation Selector-16
            || cp == 0x203C || cp == 0x2049     // ‼️, ⁉️
            || cp == 0x00A9 || cp == 0x00AE || cp == 0x2122 // ©️, ®️, ™️
            || cp == 0x3030 || cp == 0x303D || cp == 0x3297 || cp == 0x3299;
    }

    /// <summary>
    /// Splits a string into runs of plain text and individual emoji grapheme clusters.
    /// </summary>
    public static List<(string Text, bool IsEmoji)> TokenizeByEmoji(string input)
    {
        var result = new List<(string Text, bool IsEmoji)>();
        if (string.IsNullOrEmpty(input))
            return result;

        var enumerator = StringInfo.GetTextElementEnumerator(input);
        var plainBuffer = new System.Text.StringBuilder();

        while (enumerator.MoveNext())
        {
            string cluster = enumerator.GetTextElement();
            if (IsEmojiCluster(cluster))
            {
                if (plainBuffer.Length > 0)
                {
                    result.Add((plainBuffer.ToString(), false));
                    plainBuffer.Clear();
                }
                // Each emoji cluster is emitted as its own token so it can be rendered as a crisp inline image
                result.Add((cluster, true));
            }
            else
            {
                plainBuffer.Append(cluster);
            }
        }

        if (plainBuffer.Length > 0)
        {
            result.Add((plainBuffer.ToString(), false));
        }

        return result;
    }

    /// <summary>
    /// Renders an emoji grapheme cluster into a high-DPI transparent PNG in memory.
    /// Cached in RAM so each unique emoji is only rasterized once.
    /// </summary>
    public static byte[]? RenderEmojiToPng(string emojiCluster, int pixelSize = 96)
    {
        if (string.IsNullOrWhiteSpace(emojiCluster))
            return null;

        string cacheKey = $"{emojiCluster}_{pixelSize}";
        if (EmojiPngCache.TryGetValue(cacheKey, out var cached))
            return cached;

        try
        {
            var typeface = EmojiTypeface.Value;
            using var shaper = new SKShaper(typeface);
            float fontSize = pixelSize * 0.75f;
            using var font = new SKFont(typeface, fontSize)
            {
                Edging = SKFontEdging.Antialias,
                Subpixel = true
            };

            var shapeResult = shaper.Shape(emojiCluster, font);
            if (shapeResult.Points == null || shapeResult.Points.Length == 0 || shapeResult.Codepoints.All(c => c == 0))
                return null;

            int width = Math.Max(pixelSize, (int)Math.Ceiling(shapeResult.Width) + 8);
            int height = pixelSize;

            var info = new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null)
                return null;

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.Transparent);

            using var paint = new SKPaint
            {
                IsAntialias = true,
                Color = SKColors.Black
            };

            float x = Math.Max(2f, (width - shapeResult.Width) / 2f);
            float y = height * 0.78f;

            canvas.DrawShapedText(shaper, emojiCluster, x, y, SKTextAlign.Left, font, paint);
            canvas.Flush();

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            var bytes = data.ToArray();

            EmojiPngCache[cacheKey] = bytes;
            return bytes;
        }
        catch
        {
            return null;
        }
    }
}
