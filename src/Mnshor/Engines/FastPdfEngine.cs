using System.Diagnostics;
using System.IO.Compression;
using Mnshor.Models;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace Mnshor.Engines;

public class FastPdfEngine : IPdfEngine
{
    public int GetPageCount(byte[] pdfData)
    {
        if (pdfData == null || pdfData.Length == 0) return 0;
        try
        {
            using var stream = new MemoryStream(pdfData, writable: false);
            using var document = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
            return document.PageCount;
        }
        catch
        {
            return 0;
        }
    }

    public async Task<ExecutionResult> MergeAsync(IReadOnlyList<DocumentItem> documents)
    {
        if (documents == null || documents.Count < 2)
        {
            return ExecutionResult.Fail("At least two documents are required for merging.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var outputDoc = new PdfDocument();
                int totalCopiedPages = 0;

                foreach (var docItem in documents)
                {
                    using var inputStream = new MemoryStream(docItem.Data, writable: false);
                    using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import);

                    int count = inputDoc.PageCount;
                    for (int i = 0; i < count; i++)
                    {
                        outputDoc.AddPage(inputDoc.Pages[i]);
                        totalCopiedPages++;
                    }
                }

                using var outStream = new MemoryStream();
                outputDoc.Save(outStream, false);
                sw.Stop();

                string fileName = $"merged_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Merged {documents.Count} documents ({totalCopiedPages} pages) in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Merge error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> SplitRangeAsync(DocumentItem document, string rangeExpression)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import);
                int totalPages = inputDoc.PageCount;

                var pagesToExtract = ParsePageRange(rangeExpression, totalPages);
                if (pagesToExtract.Count == 0)
                {
                    return ExecutionResult.Fail($"No valid pages found in range '{rangeExpression}' (Total document pages: {totalPages}).");
                }

                using var outputDoc = new PdfDocument();
                foreach (int pageIndex in pagesToExtract)
                {
                    outputDoc.AddPage(inputDoc.Pages[pageIndex - 1]);
                }

                using var outStream = new MemoryStream();
                outputDoc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string fileName = $"{originalBase}_pages_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Extracted {pagesToExtract.Count} pages from {totalPages} pages in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Split error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> BurstToZipAsync(DocumentItem document)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import);
                int totalPages = inputDoc.PageCount;

                if (totalPages == 0)
                {
                    return ExecutionResult.Fail("Document contains no pages.");
                }

                using var zipStream = new MemoryStream();
                using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
                {
                    string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                    int padDigits = totalPages.ToString().Length;

                    for (int i = 0; i < totalPages; i++)
                    {
                        using var singlePageDoc = new PdfDocument();
                        singlePageDoc.AddPage(inputDoc.Pages[i]);

                        string entryName = $"{originalBase}_page_{(i + 1).ToString().PadLeft(padDigits, '0')}.pdf";
                        var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);

                        using var entryStream = entry.Open();
                        singlePageDoc.Save(entryStream, false);
                    }
                }

                sw.Stop();

                string originalName = Path.GetFileNameWithoutExtension(document.FileName);
                string zipFileName = $"{originalName}_burst_{DateTime.Now:yyyyMMdd_HHmmss}.zip";
                string summary = $"Burst {totalPages} pages into individual PDF files inside ZIP in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(zipStream.ToArray(), zipFileName, summary, sw.ElapsedMilliseconds, "application/zip");
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Burst error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> RotateAsync(DocumentItem document, int angleDegrees, string pageScope = "all", string? customRange = null)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        int normalizedAngle = ((angleDegrees % 360) + 360) % 360;
        if (normalizedAngle == 0)
        {
            normalizedAngle = 90;
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify);
                int totalPages = inputDoc.PageCount;

                if (totalPages == 0)
                {
                    return ExecutionResult.Fail("Document contains no pages.");
                }

                var targetPages = new HashSet<int>();
                switch (pageScope?.ToLowerInvariant())
                {
                    case "odd":
                        for (int p = 1; p <= totalPages; p += 2) targetPages.Add(p);
                        break;
                    case "even":
                        for (int p = 2; p <= totalPages; p += 2) targetPages.Add(p);
                        break;
                    case "range":
                        var parsed = ParsePageRange(customRange ?? "", totalPages);
                        if (parsed.Count == 0)
                        {
                            return ExecutionResult.Fail($"No valid pages found in range '{customRange}' (Total document pages: {totalPages}).");
                        }
                        foreach (var p in parsed) targetPages.Add(p);
                        break;
                    case "all":
                    default:
                        for (int p = 1; p <= totalPages; p++) targetPages.Add(p);
                        break;
                }

                foreach (var pageNum in targetPages)
                {
                    var page = inputDoc.Pages[pageNum - 1];
                    page.Rotate = ((page.Rotate + normalizedAngle) % 360 + 360) % 360;
                }

                using var outStream = new MemoryStream();
                inputDoc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string fileName = $"{originalBase}_rotated_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Rotated {targetPages.Count} of {totalPages} pages by {normalizedAngle}° in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Rotate error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> RotatePagesAsync(DocumentItem document, IReadOnlyList<int> pageRotations)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify);
                int totalPages = inputDoc.PageCount;

                if (totalPages == 0)
                {
                    return ExecutionResult.Fail("Document contains no pages.");
                }

                int rotatedCount = 0;
                for (int i = 0; i < totalPages; i++)
                {
                    int rot = (pageRotations != null && i < pageRotations.Count) ? pageRotations[i] : 0;
                    int norm = ((rot % 360) + 360) % 360;
                    if (norm != 0)
                    {
                        var page = inputDoc.Pages[i];
                        page.Rotate = ((page.Rotate + norm) % 360 + 360) % 360;
                        rotatedCount++;
                    }
                }

                using var outStream = new MemoryStream();
                inputDoc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string fileName = $"{originalBase}_rotated_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Rotated {rotatedCount} of {totalPages} pages in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Rotate error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> RemovePagesAsync(DocumentItem document, IReadOnlyCollection<int> pagesToRemove)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        if (pagesToRemove == null || pagesToRemove.Count == 0)
        {
            return ExecutionResult.Fail("Please select at least one page to remove.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import);
                int totalPages = inputDoc.PageCount;

                if (totalPages == 0)
                {
                    return ExecutionResult.Fail("Document contains no pages.");
                }

                var removeSet = new HashSet<int>(pagesToRemove.Where(p => p >= 1 && p <= totalPages));
                if (removeSet.Count == 0)
                {
                    return ExecutionResult.Fail("None of the specified pages exist in this document.");
                }

                if (removeSet.Count >= totalPages)
                {
                    return ExecutionResult.Fail("Cannot remove all pages. At least one page must remain in the document.");
                }

                using var outputDoc = new PdfDocument();
                int keptCount = 0;

                for (int i = 1; i <= totalPages; i++)
                {
                    if (!removeSet.Contains(i))
                    {
                        outputDoc.AddPage(inputDoc.Pages[i - 1]);
                        keptCount++;
                    }
                }

                if (keptCount == 0)
                {
                    return ExecutionResult.Fail("Cannot remove all pages. At least one page must remain.");
                }

                using var outStream = new MemoryStream();
                outputDoc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string fileName = $"{originalBase}_cleaned_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Removed {removeSet.Count} pages ({keptCount} pages remaining) in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Remove pages error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> RemovePagesAsync(DocumentItem document, string rangeExpression)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        int totalPages = GetPageCount(document.Data);
        if (totalPages == 0)
        {
            return ExecutionResult.Fail("Document contains no pages.");
        }

        var parsedPages = ParsePageRange(rangeExpression, totalPages);
        if (parsedPages.Count == 0)
        {
            return ExecutionResult.Fail($"Invalid page range '{rangeExpression}'.");
        }

        return await RemovePagesAsync(document, parsedPages);
    }

    public async Task<ExecutionResult> ReorderPagesAsync(DocumentItem document, IReadOnlyList<int> pageOrder)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        if (pageOrder == null || pageOrder.Count == 0)
        {
            return ExecutionResult.Fail("Page order list cannot be empty.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var inputDoc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import);
                int totalPages = inputDoc.PageCount;

                if (totalPages == 0)
                {
                    return ExecutionResult.Fail("Document contains no pages.");
                }

                var validPages = pageOrder.Where(p => p >= 1 && p <= totalPages).ToList();
                if (validPages.Count == 0)
                {
                    return ExecutionResult.Fail("No valid page numbers found in specified sequence.");
                }

                using var outputDoc = new PdfDocument();
                foreach (var pageNum in validPages)
                {
                    outputDoc.AddPage(inputDoc.Pages[pageNum - 1]);
                }

                using var outStream = new MemoryStream();
                outputDoc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string fileName = $"{originalBase}_reordered_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Reordered {validPages.Count} pages in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Reorder pages error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> ReversePagesAsync(DocumentItem document)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        int totalPages = GetPageCount(document.Data);
        if (totalPages == 0)
        {
            return ExecutionResult.Fail("Document contains no pages.");
        }

        var reversedOrder = Enumerable.Range(1, totalPages).Reverse().ToList();
        return await ReorderPagesAsync(document, reversedOrder);
    }

    public PdfMetadataModel GetMetadata(byte[] pdfData)
    {
        var model = new PdfMetadataModel();
        if (pdfData == null || pdfData.Length == 0) return model;

        try
        {
            using var stream = new MemoryStream(pdfData, writable: false);
            using var doc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);

            model.Title = doc.Info.Title ?? string.Empty;
            model.Author = doc.Info.Author ?? string.Empty;
            model.Subject = doc.Info.Subject ?? string.Empty;
            model.Keywords = doc.Info.Keywords ?? string.Empty;
            model.Creator = doc.Info.Creator ?? string.Empty;
            model.Producer = doc.Info.Producer ?? string.Empty;
            model.PdfVersion = $"{doc.Version / 10}.{doc.Version % 10}";

            try
            {
                if (doc.Info.CreationDate != DateTime.MinValue)
                    model.CreationDate = doc.Info.CreationDate;
            }
            catch { }

            try
            {
                if (doc.Info.ModificationDate != DateTime.MinValue)
                    model.ModificationDate = doc.Info.ModificationDate;
            }
            catch { }
        }
        catch
        {
            // Return default model if unreadable
        }

        return model;
    }

    public async Task<ExecutionResult> UpdateMetadataAsync(DocumentItem document, PdfMetadataModel metadata, bool sanitizeAll = false)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var doc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify);

                if (sanitizeAll || metadata == null)
                {
                    doc.Info.Elements.SetString("/Title", string.Empty);
                    doc.Info.Elements.SetString("/Author", string.Empty);
                    doc.Info.Elements.SetString("/Subject", string.Empty);
                    doc.Info.Elements.SetString("/Keywords", string.Empty);
                    doc.Info.Elements.SetString("/Creator", string.Empty);
                }
                else
                {
                    doc.Info.Elements.SetString("/Title", metadata.Title?.Trim() ?? string.Empty);
                    doc.Info.Elements.SetString("/Author", metadata.Author?.Trim() ?? string.Empty);
                    doc.Info.Elements.SetString("/Subject", metadata.Subject?.Trim() ?? string.Empty);
                    doc.Info.Elements.SetString("/Keywords", metadata.Keywords?.Trim() ?? string.Empty);
                    doc.Info.Elements.SetString("/Creator", metadata.Creator?.Trim() ?? string.Empty);
                }

                using var outStream = new MemoryStream();
                doc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string suffix = sanitizeAll ? "sanitized" : "metadata";
                string fileName = $"{originalBase}_{suffix}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = sanitizeAll
                    ? $"Sanitized and stripped document metadata in {sw.ElapsedMilliseconds} ms"
                    : $"Updated PDF metadata properties in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Metadata update error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> WatermarkAsync(DocumentItem document, PdfWatermarkOptions options)
    {
        if (document == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty document.");
        }

        if (options == null)
        {
            return ExecutionResult.Fail("Watermark options are required.");
        }

        if (options.Mode == WatermarkMode.Text && string.IsNullOrWhiteSpace(options.Text))
        {
            return ExecutionResult.Fail("Please enter watermark text.");
        }

        if (options.Mode == WatermarkMode.Image && (options.ImageData == null || options.ImageData.Length == 0))
        {
            return ExecutionResult.Fail("Please select a valid watermark image.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var inputStream = new MemoryStream(document.Data, writable: false);
                using var doc = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify);
                int totalPages = doc.PageCount;

                if (totalPages == 0)
                {
                    return ExecutionResult.Fail("Document contains no pages.");
                }

                var targetPages = new HashSet<int>();
                if (string.Equals(options.PageScope, "range", StringComparison.OrdinalIgnoreCase))
                {
                    var parsed = ParsePageRange(options.CustomRange ?? string.Empty, totalPages);
                    if (parsed.Count == 0)
                    {
                        return ExecutionResult.Fail($"No valid pages found in range '{options.CustomRange}' (Total document pages: {totalPages}).");
                    }
                    foreach (var p in parsed) targetPages.Add(p);
                }
                else
                {
                    for (int p = 1; p <= totalPages; p++) targetPages.Add(p);
                }

                // Render the watermark stamp once in RAM at high DPI
                var (stampPngBytes, basePointWidth, basePointHeight) = RenderWatermarkStamp(options);
                if (stampPngBytes == null || stampPngBytes.Length == 0)
                {
                    return ExecutionResult.Fail("Failed to render watermark stamp in memory.");
                }

                var layerOption = options.Layer == WatermarkLayer.BelowContent
                    ? XGraphicsPdfPageOptions.Prepend
                    : XGraphicsPdfPageOptions.Append;

                int stampedCount = 0;
                using var imgStream = new MemoryStream(stampPngBytes, 0, stampPngBytes.Length, writable: false, publiclyVisible: true);
                using var xImage = XImage.FromStream(imgStream);

                foreach (int pageNum in targetPages.OrderBy(p => p))
                {
                    var page = doc.Pages[pageNum - 1];
                    double pageW = page.Width.Point;
                    double pageH = page.Height.Point;

                    double drawW = basePointWidth;
                    double drawH = basePointHeight;

                    if (options.Mode == WatermarkMode.Image)
                    {
                        double scalePct = Math.Clamp(options.ImageScalePercent, 10, 100) / 100.0;
                        double targetW = pageW * scalePct;
                        double aspect = basePointWidth > 0 ? (basePointHeight / basePointWidth) : 1.0;
                        drawW = targetW;
                        drawH = targetW * aspect;
                    }
                    else
                    {
                        // Ensure text stamp fits comfortably within page width
                        double maxAllowedW = pageW * (options.IsMosaic ? 0.36 : 0.85);
                        if (drawW > maxAllowedW && drawW > 0)
                        {
                            double scale = maxAllowedW / drawW;
                            drawW *= scale;
                            drawH *= scale;
                        }
                    }

                    using var gfx = XGraphics.FromPdfPage(page, layerOption);

                    if (options.IsMosaic)
                    {
                        DrawMosaicWatermark(gfx, xImage, pageW, pageH, drawW, drawH, options.RotationDegrees);
                    }
                    else
                    {
                        var (cx, cy) = CalculateWatermarkCenter(pageW, pageH, drawW, drawH, options.Position, options.RotationDegrees);
                        DrawSingleWatermarkAtCenter(gfx, xImage, cx, cy, drawW, drawH, options.RotationDegrees);
                    }

                    stampedCount++;
                }

                using var outStream = new MemoryStream();
                doc.Save(outStream, false);
                sw.Stop();

                string originalBase = Path.GetFileNameWithoutExtension(document.FileName);
                string fileName = $"{originalBase}_watermarked_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string modeLabel = options.Mode == WatermarkMode.Text ? $"Text (\"{options.Text.Trim()}\")" : "Image";
                string summary = $"Applied {modeLabel} watermark to {stampedCount} of {totalPages} pages in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Watermark error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    private static void DrawSingleWatermarkAtCenter(XGraphics gfx, XImage image, double cx, double cy, double drawW, double drawH, int rotationDegrees)
    {
        var state = gfx.Save();
        gfx.TranslateTransform(cx, cy);
        if (rotationDegrees != 0)
        {
            gfx.RotateTransform(rotationDegrees);
        }
        gfx.DrawImage(image, -drawW / 2.0, -drawH / 2.0, drawW, drawH);
        gfx.Restore(state);
    }

    private static void DrawMosaicWatermark(XGraphics gfx, XImage image, double pageW, double pageH, double drawW, double drawH, int rotationDegrees)
    {
        const int cols = 3;
        const int rows = 4;
        double cellW = pageW / cols;
        double cellH = pageH / rows;

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                double cx = (c + 0.5) * cellW;
                double cy = (r + 0.5) * cellH;
                DrawSingleWatermarkAtCenter(gfx, image, cx, cy, drawW, drawH, rotationDegrees);
            }
        }
    }

    private static (double Cx, double Cy) CalculateWatermarkCenter(
        double pageW,
        double pageH,
        double drawW,
        double drawH,
        WatermarkPosition position,
        int rotationDegrees)
    {
        const double margin = 36.0;
        double rad = rotationDegrees * Math.PI / 180.0;
        double cos = Math.Abs(Math.Cos(rad));
        double sin = Math.Abs(Math.Sin(rad));

        // Bounding box of the rotated stamp so corners never clip outside the page
        double boundW = (drawW * cos) + (drawH * sin);
        double boundH = (drawW * sin) + (drawH * cos);

        double halfW = Math.Min(pageW / 2.0, boundW / 2.0 + margin);
        double halfH = Math.Min(pageH / 2.0, boundH / 2.0 + margin);

        double leftX = halfW;
        double centerX = pageW / 2.0;
        double rightX = Math.Max(halfW, pageW - halfW);

        double topY = halfH;
        double middleY = pageH / 2.0;
        double bottomY = Math.Max(halfH, pageH - halfH);

        return position switch
        {
            WatermarkPosition.TopLeft => (leftX, topY),
            WatermarkPosition.TopCenter => (centerX, topY),
            WatermarkPosition.TopRight => (rightX, topY),
            WatermarkPosition.MiddleLeft => (leftX, middleY),
            WatermarkPosition.Center => (centerX, middleY),
            WatermarkPosition.MiddleRight => (rightX, middleY),
            WatermarkPosition.BottomLeft => (leftX, bottomY),
            WatermarkPosition.BottomCenter => (centerX, bottomY),
            WatermarkPosition.BottomRight => (rightX, bottomY),
            _ => (centerX, middleY)
        };
    }

    private static (byte[]? PngBytes, double PointWidth, double PointHeight) RenderWatermarkStamp(PdfWatermarkOptions options)
    {
        byte alpha = (byte)Math.Clamp((int)Math.Round((options.OpacityPercent / 100.0) * 255), 15, 255);

        if (options.Mode == WatermarkMode.Image)
        {
            if (options.ImageData == null || options.ImageData.Length == 0)
                return (null, 0, 0);

            using var srcBitmap = SKBitmap.Decode(options.ImageData);
            if (srcBitmap == null || srcBitmap.Width <= 0 || srcBitmap.Height <= 0)
                return (null, 0, 0);

            var info = new SKImageInfo(srcBitmap.Width, srcBitmap.Height, SKColorType.Rgba8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return (null, 0, 0);

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.Transparent);

            using var paint = new SKPaint
            {
                IsAntialias = true,
                Color = new SKColor(255, 255, 255, alpha)
            };

            canvas.DrawBitmap(srcBitmap, 0, 0, paint);
            using var image = surface.Snapshot();
            using var encoded = image.Encode(SKEncodedImageFormat.Png, 100);
            return (encoded.ToArray(), srcBitmap.Width, srcBitmap.Height);
        }
        else
        {
            const float renderScale = 2.5f;
            string text = options.Text.Trim();
            float fontSizePt = Math.Clamp(options.FontSize, 12f, 140f);
            float scaledFontSize = fontSizePt * renderScale;

            using var typeface = ResolveWatermarkTypeface(options.FontFamily, options.IsBold, options.IsItalic);
            using var font = new SKFont(typeface, scaledFontSize)
            {
                Edging = SKFontEdging.SubpixelAntialias,
                Subpixel = true
            };
            using var shaper = new SKShaper(typeface);

            if (!SKColor.TryParse(options.TextColorHex, out var baseColor))
            {
                baseColor = new SKColor(239, 68, 68);
            }
            var finalColor = baseColor.WithAlpha(alpha);

            using var paint = new SKPaint
            {
                IsAntialias = true,
                Color = finalColor
            };

            bool isRtl = ContainsArabicText(text);
            var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            float spaceW = font.MeasureText(" ");
            float emojiAdvance = scaledFontSize * 1.15f;

            var measuredWords = new List<(string Word, float Width)>();
            float totalWidth = 0f;

            foreach (var word in words)
            {
                var tokens = EmojiRenderer.TokenizeByEmoji(word);
                float w = 0f;
                foreach (var (tokenText, isEmoji) in tokens)
                {
                    if (isEmoji)
                        w += emojiAdvance;
                    else
                        w += shaper.Shape(tokenText, font).Width;
                }
                measuredWords.Add((word, w));
                totalWidth += (measuredWords.Count == 1 ? w : spaceW + w);
            }

            float padX = scaledFontSize * 0.45f;
            float padY = scaledFontSize * 0.45f;
            int bmpW = Math.Max(32, (int)Math.Ceiling(totalWidth + padX * 2));
            int bmpH = Math.Max(32, (int)Math.Ceiling(scaledFontSize * 1.45f + padY * 2));

            var info = new SKImageInfo(bmpW, bmpH, SKColorType.Rgba8888, SKAlphaType.Premul);
            using var surface = SKSurface.Create(info);
            if (surface == null) return (null, 0, 0);

            var canvas = surface.Canvas;
            canvas.Clear(SKColors.Transparent);

            float baselineY = padY + scaledFontSize * 0.92f;
            float startX = (bmpW - totalWidth) / 2f;

            if (isRtl)
            {
                float cursorRight = startX + totalWidth;
                foreach (var (word, wordW) in measuredWords)
                {
                    float wordLeft = cursorRight - wordW;
                    DrawWatermarkWordTokens(canvas, word, shaper, font, paint, scaledFontSize, alpha, isRtl: true, wordLeft, wordW, baselineY);
                    cursorRight = wordLeft - spaceW;
                }
            }
            else
            {
                float cursorLeft = startX;
                foreach (var (word, wordW) in measuredWords)
                {
                    DrawWatermarkWordTokens(canvas, word, shaper, font, paint, scaledFontSize, alpha, isRtl: false, cursorLeft, wordW, baselineY);
                    cursorLeft += wordW + spaceW;
                }
            }

            using var image = surface.Snapshot();
            using var encoded = image.Encode(SKEncodedImageFormat.Png, 100);
            return (encoded.ToArray(), bmpW / renderScale, bmpH / renderScale);
        }
    }

    private static void DrawWatermarkWordTokens(
        SKCanvas canvas,
        string word,
        SKShaper shaper,
        SKFont font,
        SKPaint paint,
        float scaledFontSize,
        byte alpha,
        bool isRtl,
        float wordLeft,
        float wordW,
        float baselineY)
    {
        var tokens = EmojiRenderer.TokenizeByEmoji(word);
        float emojiBoxSize = scaledFontSize * 1.1f;
        float emojiAdvance = scaledFontSize * 1.15f;

        if (isRtl && tokens.Count > 1)
        {
            float tokenRight = wordLeft + wordW;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    float x = tokenRight - emojiAdvance;
                    DrawStampEmoji(canvas, tokenText, x, baselineY - (scaledFontSize * 0.85f), emojiBoxSize, alpha);
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
                    DrawStampEmoji(canvas, tokenText, tokenLeft, baselineY - (scaledFontSize * 0.85f), emojiBoxSize, alpha);
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

    private static void DrawStampEmoji(SKCanvas canvas, string emojiCluster, float x, float y, float size, byte alpha)
    {
        var pngBytes = EmojiRenderer.RenderEmojiToPng(emojiCluster, 96);
        if (pngBytes == null || pngBytes.Length == 0) return;

        using var skImage = SKImage.FromEncodedData(pngBytes);
        if (skImage == null) return;

        using var emojiPaint = new SKPaint
        {
            IsAntialias = true,
            Color = new SKColor(255, 255, 255, alpha)
        };

        var destRect = new SKRect(x, y, x + size, y + size);
        canvas.DrawImage(skImage, destRect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear), emojiPaint);
    }

    private static SKTypeface ResolveWatermarkTypeface(string? fontFamily, bool isBold, bool isItalic)
    {
        string fontsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string family = (fontFamily ?? "Segoe UI").Trim();

        string candidateFile = (family.ToLowerInvariant(), isBold, isItalic) switch
        {
            ("arial", true, true) => "arialbi.ttf",
            ("arial", true, false) => "arialbd.ttf",
            ("arial", false, true) => "ariali.ttf",
            ("arial", false, false) => "arial.ttf",
            ("times new roman", true, true) => "timesbi.ttf",
            ("times new roman", true, false) => "timesbd.ttf",
            ("times new roman", false, true) => "timesi.ttf",
            ("times new roman", false, false) => "times.ttf",
            ("calibri", true, true) => "calibriz.ttf",
            ("calibri", true, false) => "calibrib.ttf",
            ("calibri", false, true) => "calibrii.ttf",
            ("calibri", false, false) => "calibri.ttf",
            (_, true, true) => "segoeuiz.ttf",
            (_, true, false) => "segoeuib.ttf",
            (_, false, true) => "segoeuii.ttf",
            _ => "segoeui.ttf"
        };

        string fullPath = Path.Combine(fontsDir, candidateFile);
        if (File.Exists(fullPath))
        {
            var tf = SKTypeface.FromFile(fullPath);
            if (tf != null) return tf;
        }

        string fallbackPath = Path.Combine(fontsDir, isBold ? "arialbd.ttf" : "arial.ttf");
        if (File.Exists(fallbackPath))
        {
            var tf = SKTypeface.FromFile(fallbackPath);
            if (tf != null) return tf;
        }

        var style = new SKFontStyle(
            isBold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            isItalic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);

        return SKTypeface.FromFamilyName(family, style) ?? SKTypeface.Default;
    }

    private static bool ContainsArabicText(string text)
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

    private static List<int> ParsePageRange(string rangeExpression, int totalPages)
    {
        var result = new List<int>();
        if (string.IsNullOrWhiteSpace(rangeExpression)) return result;

        var parts = rangeExpression.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var part in parts)
        {
            if (part.Contains('-'))
            {
                var rangeLimits = part.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (rangeLimits.Length == 2 && int.TryParse(rangeLimits[0], out int start) && int.TryParse(rangeLimits[1], out int end))
                {
                    int step = start <= end ? 1 : -1;
                    for (int p = start; start <= end ? p <= end : p >= end; p += step)
                    {
                        if (p >= 1 && p <= totalPages && !result.Contains(p))
                        {
                            result.Add(p);
                        }
                    }
                }
            }
            else if (int.TryParse(part, out int singlePage))
            {
                if (singlePage >= 1 && singlePage <= totalPages && !result.Contains(singlePage))
                {
                    result.Add(singlePage);
                }
            }
        }

        return result;
    }
}
