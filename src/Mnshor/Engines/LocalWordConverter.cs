using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Mnshor.Models;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace Mnshor.Engines;

public class LocalWordConverter : IWordConverter
{
    private const float PageWidth = 595f;  // A4 width in points (72 dpi)
    private const float PageHeight = 842f; // A4 height in points (72 dpi)
    private const float MarginLeft = 42f;
    private const float MarginRight = 42f;
    private const float MarginTop = 45f;
    private const float MarginBottom = 45f;
    private const float ContentWidth = PageWidth - MarginLeft - MarginRight;

    private static readonly Lazy<SKTypeface> RegularTypeface = new(() => LoadSystemTypeface("segoeui.ttf", "arial.ttf"));
    private static readonly Lazy<SKTypeface> BoldTypeface = new(() => LoadSystemTypeface("segoeuib.ttf", "arialbd.ttf"));

    private static SKTypeface LoadSystemTypeface(string primaryFile, string fallbackFile)
    {
        string fontsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string p1 = Path.Combine(fontsDir, primaryFile);
        if (File.Exists(p1))
        {
            var tf = SKTypeface.FromFile(p1);
            if (tf != null) return tf;
        }

        string p2 = Path.Combine(fontsDir, fallbackFile);
        if (File.Exists(p2))
        {
            var tf = SKTypeface.FromFile(p2);
            if (tf != null) return tf;
        }

        return SKTypeface.FromFamilyName("Segoe UI") ?? SKTypeface.FromFamilyName("Arial") ?? SKTypeface.Default;
    }

    public async Task<ExecutionResult> ConvertToPdfAsync(DocumentItem document)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty Word document.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var docxStream = new MemoryStream(document.Data, writable: false);
                using var wordDoc = WordprocessingDocument.Open(docxStream, false);

                var mainPart = wordDoc.MainDocumentPart;
                var body = mainPart?.Document?.Body;
                if (body == null)
                {
                    return ExecutionResult.Fail("Word document has no readable body.");
                }

                using var outStream = new MemoryStream();
                var metadata = new SKDocumentPdfMetadata
                {
                    Title = Path.GetFileNameWithoutExtension(document.FileName),
                    Creator = "Mnshor Local PDF Suite",
                    Producer = "Mnshor High-Performance Engine",
                    RasterDpi = 300f
                };

                using (var pdfDoc = SKDocument.CreatePdf(outStream, metadata))
                {
                    using var ctx = new SkiaLayoutContext(pdfDoc);

                    foreach (var element in body.Elements())
                    {
                        if (element is Paragraph para)
                        {
                            RenderParagraph(ctx, mainPart!, para);
                        }
                        else if (element is Table table)
                        {
                            RenderTable(ctx, mainPart!, table);
                        }
                    }

                    ctx.Finish();
                    int totalPages = ctx.PageCount;
                    pdfDoc.Close();

                    sw.Stop();

                    string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                    string fileName = $"{originalBase}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                    string summary = $"Converted Word document ({totalPages} pages) in {sw.ElapsedMilliseconds} ms";

                    return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Conversion error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    private void RenderParagraph(SkiaLayoutContext ctx, MainDocumentPart mainPart, Paragraph para)
    {
        // 1. Check for embedded images (<w:drawing>) inside the paragraph
        var blipEmbeds = para.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
            .Select(b => b.Embed?.Value)
            .Where(id => !string.IsNullOrEmpty(id))
            .ToList();

        foreach (var embedId in blipEmbeds)
        {
            RenderEmbeddedImage(ctx, mainPart, embedId!);
        }

        // 2. Extract full paragraph text (including nested runs, hyperlinks, content controls, line breaks)
        string fullText = ExtractParagraphText(para);
        if (string.IsNullOrWhiteSpace(fullText))
        {
            if (blipEmbeds.Count == 0)
            {
                ctx.AdvanceY(10f);
            }
            return;
        }

        // 3. Determine styling (Heading, Bold, Font Size, Color, List Bullet, Alignment, RTL)
        string? styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        bool isHeading1 = styleId != null && (styleId.Contains("Heading1", StringComparison.OrdinalIgnoreCase) || styleId.Contains("Title", StringComparison.OrdinalIgnoreCase));
        bool isHeading2 = styleId != null && (styleId.Contains("Heading2", StringComparison.OrdinalIgnoreCase) || styleId.Contains("Subtitle", StringComparison.OrdinalIgnoreCase));
        bool isHeading = isHeading1 || isHeading2 || (styleId != null && styleId.Contains("Heading", StringComparison.OrdinalIgnoreCase));

        bool hasBoldRun = isHeading || para.Descendants<RunProperties>().Any(rp => rp.Bold != null && (rp.Bold.Val == null || rp.Bold.Val.Value));
        float fontSize = isHeading1 ? 18f : isHeading2 ? 15f : isHeading ? 13.5f : 11.5f;

        // Check explicit font size in half-points (<w:sz w:val="28"/> -> 14pt)
        var szVal = para.Descendants<FontSize>().FirstOrDefault()?.Val?.Value;
        if (!string.IsNullOrEmpty(szVal) && float.TryParse(szVal, out float halfPts) && halfPts > 8 && halfPts < 144)
        {
            fontSize = halfPts / 2f;
        }

        // Check list item (<w:numPr>)
        bool isListItem = para.ParagraphProperties?.NumberingProperties != null;
        if (isListItem && !fullText.TrimStart().StartsWith("•"))
        {
            fullText = "• " + fullText;
        }

        // Check RTL (Arabic text or explicit BiDi property)
        bool isRtl = ContainsArabic(fullText) || para.ParagraphProperties?.BiDi != null;

        // Check text color
        SKColor textColor = isHeading1 ? new SKColor(15, 23, 42) : new SKColor(30, 41, 59);
        var colorVal = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Color>().FirstOrDefault()?.Val?.Value;
        if (!string.IsNullOrEmpty(colorVal) && colorVal.Length == 6 && !colorVal.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            if (SKColor.TryParse("#" + colorVal, out var parsedColor))
            {
                textColor = parsedColor;
            }
        }

        var typeface = hasBoldRun ? BoldTypeface.Value : RegularTypeface.Value;
        float lineHeight = fontSize * 1.55f;

        // Split by explicit line breaks (\n)
        var logicalLines = fullText.Split('\n');
        foreach (var logicalLine in logicalLines)
        {
            if (string.IsNullOrWhiteSpace(logicalLine))
            {
                ctx.AdvanceY(lineHeight * 0.6f);
                continue;
            }

            RenderWrappedLine(ctx, logicalLine.Trim(), typeface, fontSize, lineHeight, textColor, isRtl, MarginLeft, ContentWidth);
        }

        ctx.AdvanceY(isHeading ? 8f : 4f);
    }

    private void RenderEmbeddedImage(SkiaLayoutContext ctx, MainDocumentPart mainPart, string relationshipId)
    {
        try
        {
            if (mainPart.GetPartById(relationshipId) is ImagePart imagePart)
            {
                using var stream = imagePart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                using var skImage = SKImage.FromEncodedData(ms.ToArray());
                if (skImage != null)
                {
                    float maxW = Math.Min(ContentWidth, 360f);
                    float scale = Math.Min(1f, maxW / skImage.Width);
                    float drawW = skImage.Width * scale;
                    float drawH = skImage.Height * scale;

                    ctx.EnsureSpace(drawH + 12f);
                    var destRect = new SKRect(MarginLeft, ctx.CurrentY, MarginLeft + drawW, ctx.CurrentY + drawH);
                    ctx.Canvas.DrawImage(skImage, destRect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
                    ctx.AdvanceY(drawH + 10f);
                }
            }
        }
        catch
        {
            // Ignore corrupted embedded image parts
        }
    }

    private void RenderWrappedLine(
        SkiaLayoutContext ctx,
        string text,
        SKTypeface typeface,
        float fontSize,
        float lineHeight,
        SKColor textColor,
        bool isRtl,
        float startX,
        float maxWidth)
    {
        using var font = new SKFont(typeface, fontSize)
        {
            Edging = SKFontEdging.SubpixelAntialias,
            Subpixel = true
        };
        using var shaper = new SKShaper(typeface);
        using var paint = new SKPaint
        {
            IsAntialias = true,
            Color = textColor
        };

        var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0) return;

        float spaceWidth = font.MeasureText(" ");
        var currentLineWords = new List<(string Word, float Width)>();
        float currentLineWidth = 0f;

        foreach (var word in words)
        {
            float wordWidth = MeasureWordWithEmojis(word, shaper, font, fontSize);
            float projectedWidth = currentLineWords.Count == 0 ? wordWidth : currentLineWidth + spaceWidth + wordWidth;

            if (projectedWidth > maxWidth && currentLineWords.Count > 0)
            {
                ctx.EnsureSpace(lineHeight);
                DrawVisualLine(ctx.Canvas, currentLineWords, shaper, font, paint, fontSize, isRtl, startX, maxWidth, spaceWidth, ctx.CurrentY + fontSize);
                ctx.AdvanceY(lineHeight);

                currentLineWords.Clear();
                currentLineWords.Add((word, wordWidth));
                currentLineWidth = wordWidth;
            }
            else
            {
                currentLineWords.Add((word, wordWidth));
                currentLineWidth = projectedWidth;
            }
        }

        if (currentLineWords.Count > 0)
        {
            ctx.EnsureSpace(lineHeight);
            DrawVisualLine(ctx.Canvas, currentLineWords, shaper, font, paint, fontSize, isRtl, startX, maxWidth, spaceWidth, ctx.CurrentY + fontSize);
            ctx.AdvanceY(lineHeight);
        }
    }

    private static float MeasureWordWithEmojis(string word, SKShaper shaper, SKFont font, float fontSize)
    {
        var tokens = EmojiRenderer.TokenizeByEmoji(word);
        float width = 0f;
        float emojiAdvance = fontSize * 1.15f;

        foreach (var (tokenText, isEmoji) in tokens)
        {
            if (isEmoji)
            {
                width += emojiAdvance;
            }
            else
            {
                var shaped = shaper.Shape(tokenText, font);
                width += shaped.Width;
            }
        }

        return width;
    }

    private static void DrawVisualLine(
        SKCanvas canvas,
        List<(string Word, float Width)> words,
        SKShaper shaper,
        SKFont font,
        SKPaint paint,
        float fontSize,
        bool isRtl,
        float startX,
        float maxWidth,
        float spaceWidth,
        float baselineY)
    {
        if (isRtl)
        {
            // Right-to-Left layout: start from right edge and flow leftward
            float cursorRight = startX + maxWidth;
            foreach (var (word, wordWidth) in words)
            {
                float wordLeft = cursorRight - wordWidth;
                DrawWordTokens(canvas, word, shaper, font, paint, fontSize, isRtl: true, wordLeft, wordWidth, baselineY);
                cursorRight = wordLeft - spaceWidth;
            }
        }
        else
        {
            // Left-to-Right layout: start from left edge and flow rightward
            float cursorLeft = startX;
            foreach (var (word, wordWidth) in words)
            {
                DrawWordTokens(canvas, word, shaper, font, paint, fontSize, isRtl: false, cursorLeft, wordWidth, baselineY);
                cursorLeft += wordWidth + spaceWidth;
            }
        }
    }

    private static void DrawWordTokens(
        SKCanvas canvas,
        string word,
        SKShaper shaper,
        SKFont font,
        SKPaint paint,
        float fontSize,
        bool isRtl,
        float wordLeft,
        float wordWidth,
        float baselineY)
    {
        var tokens = EmojiRenderer.TokenizeByEmoji(word);
        float emojiBoxSize = fontSize * 1.12f;
        float emojiAdvance = fontSize * 1.15f;

        if (isRtl && tokens.Count > 1)
        {
            // In RTL, tokens within a compound word flow right-to-left
            float tokenRight = wordLeft + wordWidth;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    float x = tokenRight - emojiAdvance;
                    DrawEmojiImage(canvas, tokenText, x, baselineY - (fontSize * 0.85f), emojiBoxSize);
                    tokenRight -= emojiAdvance;
                }
                else
                {
                    var shaped = shaper.Shape(tokenText, font);
                    float x = tokenRight - shaped.Width;
                    canvas.DrawShapedText(shaper, tokenText, x, baselineY, SKTextAlign.Left, font, paint);
                    tokenRight -= shaped.Width;
                }
            }
        }
        else
        {
            float tokenLeft = wordLeft;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    DrawEmojiImage(canvas, tokenText, tokenLeft, baselineY - (fontSize * 0.85f), emojiBoxSize);
                    tokenLeft += emojiAdvance;
                }
                else
                {
                    var shaped = shaper.Shape(tokenText, font);
                    canvas.DrawShapedText(shaper, tokenText, tokenLeft, baselineY, SKTextAlign.Left, font, paint);
                    tokenLeft += shaped.Width;
                }
            }
        }
    }

    private static void DrawEmojiImage(SKCanvas canvas, string emojiCluster, float x, float y, float size)
    {
        var pngBytes = EmojiRenderer.RenderEmojiToPng(emojiCluster, 96);
        if (pngBytes == null || pngBytes.Length == 0)
            return;

        using var skImage = SKImage.FromEncodedData(pngBytes);
        if (skImage == null)
            return;

        var destRect = new SKRect(x, y, x + size, y + size);
        canvas.DrawImage(skImage, destRect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
    }

    private void RenderTable(SkiaLayoutContext ctx, MainDocumentPart mainPart, Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0) return;

        ctx.EnsureSpace(32f);
        ctx.AdvanceY(6f);

        using var borderPaint = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            Color = new SKColor(203, 213, 225),
            StrokeWidth = 0.8f,
            IsAntialias = true
        };

        using var headerFillPaint = new SKPaint
        {
            Style = SKPaintStyle.Fill,
            Color = new SKColor(241, 245, 249),
            IsAntialias = true
        };

        bool isFirstRow = true;

        foreach (var row in rows)
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count == 0) continue;

            float colWidth = ContentWidth / cells.Count;
            float rowHeight = 26f;

            ctx.EnsureSpace(rowHeight);

            for (int i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                // In RTL tables, order columns from right to left if content is Arabic
                float x = MarginLeft + (i * colWidth);
                float y = ctx.CurrentY;
                var cellRect = new SKRect(x, y, x + colWidth, y + rowHeight);

                if (isFirstRow)
                {
                    ctx.Canvas.DrawRect(cellRect, headerFillPaint);
                }

                ctx.Canvas.DrawRect(cellRect, borderPaint);

                string cellText = cell.InnerText.Trim();
                if (!string.IsNullOrEmpty(cellText))
                {
                    var tf = isFirstRow ? BoldTypeface.Value : RegularTypeface.Value;
                    float fontSize = 10f;
                    using var font = new SKFont(tf, fontSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
                    using var shaper = new SKShaper(tf);
                    using var textPaint = new SKPaint { IsAntialias = true, Color = new SKColor(15, 23, 42) };

                    bool cellRtl = ContainsArabic(cellText);
                    var words = cellText.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                        .Select(w => (w, MeasureWordWithEmojis(w, shaper, font, fontSize)))
                        .ToList();

                    float spaceW = font.MeasureText(" ");
                    DrawVisualLine(ctx.Canvas, words, shaper, font, textPaint, fontSize, cellRtl, x + 6f, colWidth - 12f, spaceW, y + 17f);
                }
            }

            ctx.AdvanceY(rowHeight);
            isFirstRow = false;
        }

        ctx.AdvanceY(10f);
    }

    private static string ExtractParagraphText(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.Descendants<OpenXmlElement>())
        {
            if (child is Text t)
            {
                sb.Append(t.Text);
            }
            else if (child is Break)
            {
                sb.Append('\n');
            }
            else if (child is TabChar)
            {
                sb.Append("    ");
            }
            else if (child is NoBreakHyphen)
            {
                sb.Append('-');
            }
        }
        return sb.ToString();
    }

    private static bool ContainsArabic(string text)
    {
        foreach (char c in text)
        {
            if ((c >= '\u0600' && c <= '\u06FF') ||
                (c >= '\u0750' && c <= '\u077F') ||
                (c >= '\uFB50' && c <= '\uFDFF') ||
                (c >= '\uFE70' && c <= '\uFEFF'))
            {
                return true;
            }
        }
        return false;
    }

    private sealed class SkiaLayoutContext : IDisposable
    {
        private readonly SKDocument _pdfDoc;
        public SKCanvas Canvas { get; private set; } = null!;
        public float CurrentY { get; private set; }
        public int PageCount { get; private set; }
        private bool _pageOpen;

        public SkiaLayoutContext(SKDocument pdfDoc)
        {
            _pdfDoc = pdfDoc;
            StartNewPage();
        }

        public void StartNewPage()
        {
            if (_pageOpen)
            {
                _pdfDoc.EndPage();
                _pageOpen = false;
            }

            Canvas = _pdfDoc.BeginPage(PageWidth, PageHeight);
            Canvas.Clear(SKColors.White);
            CurrentY = MarginTop;
            PageCount++;
            _pageOpen = true;
        }

        public void AdvanceY(float delta)
        {
            CurrentY += delta;
        }

        public void EnsureSpace(float requiredHeight)
        {
            if (CurrentY + requiredHeight > (PageHeight - MarginBottom))
            {
                StartNewPage();
            }
        }

        public void Finish()
        {
            if (_pageOpen)
            {
                _pdfDoc.EndPage();
                _pageOpen = false;
            }
        }

        public void Dispose()
        {
            Finish();
        }
    }
}
