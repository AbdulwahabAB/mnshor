using System.Collections.Concurrent;
using PdfSharp.Fonts;

namespace Mnshor.Engines;

public class SystemFontResolver : IFontResolver
{
    private static readonly ConcurrentDictionary<string, byte[]> FontCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly string WindowsFontsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

    private static readonly Dictionary<string, string> FontMappings = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Segoe UI Emoji"] = "seguiemj.ttf",
        ["Segoe UI"] = "segoeui.ttf",
        ["Arial"] = "arial.ttf",
        ["Calibri"] = "calibri.ttf",
        ["Times New Roman"] = "times.ttf"
    };

    public byte[]? GetFont(string faceName)
    {
        if (FontCache.TryGetValue(faceName, out var cachedBytes))
        {
            return cachedBytes;
        }

        string fileName = faceName.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ? faceName : $"{faceName}.ttf";
        string filePath = Path.Combine(WindowsFontsDir, fileName);

        if (File.Exists(filePath))
        {
            try
            {
                var bytes = File.ReadAllBytes(filePath);
                FontCache[faceName] = bytes;
                return bytes;
            }
            catch
            {
                return null;
            }
        }

        return null;
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (FontMappings.TryGetValue(familyName, out var fontFile))
        {
            return new FontResolverInfo(fontFile);
        }

        // Generic fallback to Segoe UI or Arial
        if (familyName.Contains("Emoji", StringComparison.OrdinalIgnoreCase))
        {
            return new FontResolverInfo("seguiemj.ttf");
        }

        return new FontResolverInfo("arial.ttf");
    }
}
