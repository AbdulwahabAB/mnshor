using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Mnshor.Models;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace Mnshor.Engines;

public class LocalSpreadsheetMarkdownConverter : ISpreadsheetMarkdownConverter
{
    private static readonly Lazy<SKTypeface> RegularTypeface = new(() => LoadSystemTypeface("segoeui.ttf", "arial.ttf", "Segoe UI", SKFontStyle.Normal));
    private static readonly Lazy<SKTypeface> BoldTypeface = new(() => LoadSystemTypeface("segoeuib.ttf", "arialbd.ttf", "Segoe UI", SKFontStyle.Bold));
    private static readonly Lazy<SKTypeface> ItalicTypeface = new(() => LoadSystemTypeface("segoeuii.ttf", "ariali.ttf", "Segoe UI", SKFontStyle.Italic));
    private static readonly Lazy<SKTypeface> BoldItalicTypeface = new(() => LoadSystemTypeface("segoeuiz.ttf", "arialbi.ttf", "Segoe UI", SKFontStyle.BoldItalic));
    private static readonly Lazy<SKTypeface> MonoTypeface = new(() => LoadSystemTypeface("consola.ttf", "cour.ttf", "Consolas", SKFontStyle.Normal));

    private static SKTypeface LoadSystemTypeface(string primaryFile, string fallbackFile, string familyName, SKFontStyle style)
    {
        try
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
        }
        catch
        {
            // Fallback to family name lookup
        }

        return SKTypeface.FromFamilyName(familyName, style)
            ?? SKTypeface.FromFamilyName("Segoe UI", style)
            ?? SKTypeface.FromFamilyName("Arial", style)
            ?? SKTypeface.Default;
    }

    public ExcelWorkbookInfo InspectWorkbook(byte[] xlsxData)
    {
        if (xlsxData == null || xlsxData.Length == 0)
        {
            return new ExcelWorkbookInfo
            {
                IsValid = false,
                ErrorMessage = "Invalid or empty Excel (.xlsx) file."
            };
        }

        try
        {
            var sheets = ExtractWorksheets(xlsxData);
            if (sheets.Count == 0)
            {
                return new ExcelWorkbookInfo
                {
                    IsValid = false,
                    ErrorMessage = "Excel workbook contains no worksheets."
                };
            }

            var info = new ExcelWorkbookInfo { IsValid = true };
            foreach (var sheet in sheets)
            {
                info.Sheets.Add(new ExcelSheetPreview
                {
                    Name = sheet.Name,
                    TotalRows = sheet.Rows.Count,
                    TotalColumns = sheet.ColumnCount,
                    IsRtlDetected = sheet.IsRtl,
                    PreviewRows = sheet.Rows.Take(25).ToList()
                });
            }

            return info;
        }
        catch (Exception ex)
        {
            return new ExcelWorkbookInfo
            {
                IsValid = false,
                ErrorMessage = $"Failed to read Excel workbook: {ex.Message}"
            };
        }
    }

    public async Task<ExecutionResult> ConvertExcelToPdfAsync(DocumentItem document, SpreadsheetMarkdownOptions? options = null)
    {
        if (document == null || document.Data == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty Excel (.xlsx) document.");
        }

        options ??= new SpreadsheetMarkdownOptions
        {
            Mode = ConversionInputMode.Excel,
            Orientation = PdfPageOrientation.Landscape
        };

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                var allSheets = ExtractWorksheets(document.Data);
                if (allSheets.Count == 0)
                {
                    return ExecutionResult.Fail("No worksheets found in the Excel workbook.");
                }

                var targetSheets = allSheets
                    .Where(s => options.SelectedSheetNames.Count == 0 || options.SelectedSheetNames.Contains(s.Name))
                    .ToList();

                if (targetSheets.Count == 0)
                {
                    return ExecutionResult.Fail("Please select at least one worksheet to convert.");
                }

                // Filter to sheets that have data, or if all selected are empty keep them so user gets a valid page
                var nonEmptySheets = targetSheets.Where(s => s.Rows.Count > 0 && s.ColumnCount > 0).ToList();
                if (nonEmptySheets.Count == 0)
                {
                    return ExecutionResult.Fail("The selected worksheet(s) do not contain any cell data.");
                }

                var (pageW, pageH) = GetPageDimensions(options.PageSize, options.Orientation);
                SKColor accentColor = ParseColorOrDefault(options.HeaderAccentHex, new SKColor(30, 41, 59));
                string docTitle = Path.GetFileNameWithoutExtension(document.FileName);

                using var outStream = new MemoryStream();
                var metadata = new SKDocumentPdfMetadata
                {
                    Title = docTitle,
                    Creator = "Mnshor Local PDF Suite",
                    Producer = "Mnshor High-Performance Engine",
                    RasterDpi = 300f
                };

                using (var pdfDoc = SKDocument.CreatePdf(outStream, metadata))
                {
                    using var ctx = new PaginatedLayoutContext(pdfDoc, pageW, pageH, marginH: 34f, marginV: 38f, docTitle, options.IncludePageNumbers);

                    for (int sIdx = 0; sIdx < nonEmptySheets.Count; sIdx++)
                    {
                        var sheet = nonEmptySheets[sIdx];
                        if (sIdx > 0 && ctx.CurrentY > ctx.MarginTop + 10f)
                        {
                            ctx.StartNewPage();
                        }

                        bool useRtl = options.DirectionMode switch
                        {
                            TableDirectionMode.Rtl => true,
                            TableDirectionMode.Ltr => false,
                            _ => sheet.IsRtl
                        };

                        RenderExcelSheet(ctx, sheet, options, accentColor, useRtl);
                    }

                    ctx.Finish();
                    int totalPages = ctx.PageCount;
                    pdfDoc.Close();

                    sw.Stop();
                    string fileName = $"{docTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                    int totalRowsRendered = nonEmptySheets.Sum(s => s.Rows.Count);
                    string summary = $"Converted {nonEmptySheets.Count} worksheet(s) ({totalRowsRendered} rows → {totalPages} PDF page(s)) in {sw.ElapsedMilliseconds} ms";

                    return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Excel conversion error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    public async Task<ExecutionResult> ConvertMarkdownToPdfAsync(string markdownText, string documentTitle, SpreadsheetMarkdownOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(markdownText))
        {
            return ExecutionResult.Fail("Markdown content is empty. Please provide Markdown text or upload a .md file.");
        }

        options ??= new SpreadsheetMarkdownOptions
        {
            Mode = ConversionInputMode.Markdown,
            Orientation = PdfPageOrientation.Portrait
        };

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                var blocks = ParseMarkdownBlocks(markdownText);
                if (blocks.Count == 0)
                {
                    return ExecutionResult.Fail("No readable Markdown blocks found.");
                }

                string cleanTitle = string.IsNullOrWhiteSpace(documentTitle)
                    ? "Markdown_Document"
                    : Path.GetFileNameWithoutExtension(documentTitle.Trim());

                var (pageW, pageH) = GetPageDimensions(options.PageSize, options.Orientation);
                SKColor accentColor = ParseColorOrDefault(options.HeaderAccentHex, new SKColor(37, 99, 235));
                float scaleFactor = options.FontSizeScale switch
                {
                    MarkdownFontSizeScale.Small => 0.88f,
                    MarkdownFontSizeScale.Large => 1.14f,
                    _ => 1.0f
                };

                using var outStream = new MemoryStream();
                var metadata = new SKDocumentPdfMetadata
                {
                    Title = cleanTitle,
                    Creator = "Mnshor Local PDF Suite",
                    Producer = "Mnshor High-Performance Engine",
                    RasterDpi = 300f
                };

                using (var pdfDoc = SKDocument.CreatePdf(outStream, metadata))
                {
                    using var ctx = new PaginatedLayoutContext(pdfDoc, pageW, pageH, marginH: 44f, marginV: 44f, cleanTitle, options.IncludePageNumbers);

                    foreach (var block in blocks)
                    {
                        RenderMarkdownBlock(ctx, block, scaleFactor, accentColor, options);
                    }

                    ctx.Finish();
                    int totalPages = ctx.PageCount;
                    pdfDoc.Close();

                    sw.Stop();
                    string fileName = $"{cleanTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                    string summary = $"Converted Markdown ({blocks.Count} blocks → {totalPages} PDF page(s)) in {sw.ElapsedMilliseconds} ms";

                    return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Markdown conversion error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    // =========================================================================
    // 1. OpenXML Excel (.xlsx) Parsing
    // =========================================================================

    private sealed class ExtractedWorksheet
    {
        public string Name { get; set; } = string.Empty;
        public List<string[]> Rows { get; set; } = new();
        public int ColumnCount { get; set; }
        public bool IsRtl { get; set; }
    }

    private static List<ExtractedWorksheet> ExtractWorksheets(byte[] xlsxData)
    {
        using var ms = new MemoryStream(xlsxData, writable: false);
        using var doc = SpreadsheetDocument.Open(ms, isEditable: false);

        var workbookPart = doc.WorkbookPart;
        if (workbookPart?.Workbook?.Sheets == null)
        {
            return new List<ExtractedWorksheet>();
        }

        // Load Shared String Table into an indexed array for O(1) lookup
        string[] sharedStrings = Array.Empty<string>();
        var sstPart = workbookPart.SharedStringTablePart;
        if (sstPart?.SharedStringTable != null)
        {
            sharedStrings = sstPart.SharedStringTable
                .Elements<SharedStringItem>()
                .Select(item => item.InnerText ?? string.Empty)
                .ToArray();
        }

        // Detect date number format style indices in Stylesheet
        var dateStyleIndices = ExtractDateStyleIndices(workbookPart.WorkbookStylesPart);

        var result = new List<ExtractedWorksheet>();

        foreach (var sheet in workbookPart.Workbook.Sheets.Elements<Sheet>())
        {
            string sheetName = sheet.Name?.Value ?? $"Sheet {result.Count + 1}";
            string? relId = sheet.Id?.Value;
            if (string.IsNullOrEmpty(relId))
                continue;

            if (workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart)
                continue;

            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null)
                continue;

            // Check explicit RightToLeft view flag in worksheet
            bool explicitRtl = worksheet.Descendants<SheetView>().Any(sv => sv.RightToLeft != null && sv.RightToLeft.Value);

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                result.Add(new ExtractedWorksheet { Name = sheetName, IsRtl = explicitRtl });
                continue;
            }

            var rawRows = new List<Dictionary<int, string>>();
            int maxColIndex = -1;
            bool hasArabicContent = false;

            foreach (var row in sheetData.Elements<Row>())
            {
                var rowDict = new Dictionary<int, string>();
                int fallbackColIndex = 0;

                foreach (var cell in row.Elements<Cell>())
                {
                    int colIdx = ParseColumnIndex(cell.CellReference?.Value, fallbackColIndex);
                    fallbackColIndex = colIdx + 1;

                    string val = GetCellText(cell, sharedStrings, dateStyleIndices).Trim();
                    if (!string.IsNullOrEmpty(val))
                    {
                        rowDict[colIdx] = val;
                        if (colIdx > maxColIndex)
                        {
                            maxColIndex = colIdx;
                        }
                        if (!hasArabicContent && ContainsArabic(val))
                        {
                            hasArabicContent = true;
                        }
                    }
                }

                // Keep row even if empty initially; we will trim completely empty leading/trailing rows
                rawRows.Add(rowDict);
            }

            // Remove completely empty rows
            var nonEmptyRowDicts = rawRows.Where(r => r.Count > 0).ToList();
            if (nonEmptyRowDicts.Count == 0 || maxColIndex < 0)
            {
                result.Add(new ExtractedWorksheet
                {
                    Name = sheetName,
                    ColumnCount = 0,
                    IsRtl = explicitRtl
                });
                continue;
            }

            // Find minimum non-empty column index so sheets starting at Column B or C don't have blank leading columns
            int minColIndex = nonEmptyRowDicts.SelectMany(r => r.Keys).Min();
            int colCount = maxColIndex - minColIndex + 1;

            var normalizedRows = new List<string[]>(nonEmptyRowDicts.Count);
            foreach (var rDict in nonEmptyRowDicts)
            {
                var rowArr = new string[colCount];
                for (int c = 0; c < colCount; c++)
                {
                    rowArr[c] = rDict.TryGetValue(minColIndex + c, out var v) ? v : string.Empty;
                }
                normalizedRows.Add(rowArr);
            }

            result.Add(new ExtractedWorksheet
            {
                Name = sheetName,
                Rows = normalizedRows,
                ColumnCount = colCount,
                IsRtl = explicitRtl || hasArabicContent
            });
        }

        return result;
    }

    private static HashSet<uint> ExtractDateStyleIndices(WorkbookStylesPart? stylesPart)
    {
        var dateFormatIds = new HashSet<uint>
        {
            14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57
        };

        var dateStyleIndices = new HashSet<uint>();
        var stylesheet = stylesPart?.Stylesheet;
        if (stylesheet == null)
            return dateStyleIndices;

        if (stylesheet.NumberingFormats != null)
        {
            foreach (var nf in stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                uint? id = nf.NumberFormatId?.Value;
                string? code = nf.FormatCode?.Value;
                if (id.HasValue && !string.IsNullOrEmpty(code))
                {
                    // Remove escaped literals and color brackets
                    string cleaned = Regex.Replace(code, @"\[[^\]]*\]|""[^""]*""", "");
                    if (cleaned.Contains('y', StringComparison.OrdinalIgnoreCase) &&
                        cleaned.Contains('d', StringComparison.OrdinalIgnoreCase))
                    {
                        dateFormatIds.Add(id.Value);
                    }
                }
            }
        }

        if (stylesheet.CellFormats != null)
        {
            uint styleIdx = 0;
            foreach (var cf in stylesheet.CellFormats.Elements<CellFormat>())
            {
                uint? numFmtId = cf.NumberFormatId?.Value;
                if (numFmtId.HasValue && dateFormatIds.Contains(numFmtId.Value))
                {
                    dateStyleIndices.Add(styleIdx);
                }
                styleIdx++;
            }
        }

        return dateStyleIndices;
    }

    private static int ParseColumnIndex(string? cellReference, int fallbackIndex)
    {
        if (string.IsNullOrEmpty(cellReference))
            return fallbackIndex;

        int colIndex = 0;
        int letterCount = 0;
        foreach (char ch in cellReference)
        {
            if (char.IsLetter(ch))
            {
                char upper = char.ToUpperInvariant(ch);
                colIndex = (colIndex * 26) + (upper - 'A' + 1);
                letterCount++;
            }
            else
            {
                break;
            }
        }

        return letterCount > 0 ? colIndex - 1 : fallbackIndex;
    }

    private static string GetCellText(Cell cell, string[] sharedStrings, HashSet<uint> dateStyleIndices)
    {
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? cell.InnerText ?? string.Empty;
        }

        string rawVal = cell.CellValue?.InnerText ?? cell.InnerText ?? string.Empty;
        if (string.IsNullOrEmpty(rawVal))
            return string.Empty;

        if (dataType == CellValues.SharedString)
        {
            if (int.TryParse(rawVal, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sstIdx) &&
                sstIdx >= 0 && sstIdx < sharedStrings.Length)
            {
                return sharedStrings[sstIdx];
            }
            return rawVal;
        }

        if (dataType == CellValues.Boolean)
        {
            return rawVal == "1" || rawVal.Equals("true", StringComparison.OrdinalIgnoreCase) ? "TRUE" : "FALSE";
        }

        if (dataType == CellValues.Date)
        {
            if (DateTime.TryParse(rawVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            return rawVal;
        }

        // Numeric or unlabelled cell value
        if ((dataType == null || dataType == CellValues.Number) &&
            double.TryParse(rawVal, NumberStyles.Float, CultureInfo.InvariantCulture, out double numVal))
        {
            uint? styleIdx = cell.StyleIndex?.Value;
            if (styleIdx.HasValue && dateStyleIndices.Contains(styleIdx.Value) && numVal > 1 && numVal < 2958465)
            {
                try
                {
                    return DateTime.FromOADate(numVal).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                }
                catch
                {
                    // Fallback to number formatting
                }
            }

            // Format clean decimals if floating point noise exists (e.g. 19.990000000000002 -> 19.99)
            if (Math.Abs(numVal - Math.Round(numVal)) < 1e-9)
            {
                return Math.Round(numVal).ToString("0", CultureInfo.InvariantCulture);
            }
            return Math.Round(numVal, 4).ToString("G", CultureInfo.InvariantCulture);
        }

        return rawVal;
    }

    // =========================================================================
    // 2. Excel Sheet Table Layout & Paginated Renderer
    // =========================================================================

    private void RenderExcelSheet(
        PaginatedLayoutContext ctx,
        ExtractedWorksheet sheet,
        SpreadsheetMarkdownOptions options,
        SKColor accentColor,
        bool isRtl)
    {
        if (options.ShowSheetTitle)
        {
            ctx.EnsureSpace(38f);
            DrawSheetTitleBanner(ctx, sheet.Name, sheet.Rows.Count, sheet.ColumnCount, accentColor, isRtl);
        }

        int dataCols = sheet.ColumnCount;
        if (dataCols <= 0 || sheet.Rows.Count == 0)
            return;

        bool showRowNum = options.ShowRowNumbers;
        float rowNumColWidth = showRowNum ? 30f : 0f;
        float availableDataWidth = ctx.ContentWidth - rowNumColWidth;

        // Calculate proportional column widths based on max character lengths per column
        float[] dataColWidths = ComputeProportionalColumnWidths(sheet.Rows, dataCols, availableDataWidth);

        // Choose font size based on column count
        float fontSize = dataCols switch
        {
            <= 5 => 10f,
            <= 8 => 9f,
            <= 12 => 8.2f,
            _ => 7.4f
        };

        float lineHeight = fontSize * 1.42f;
        float cellPadX = 5f;
        float cellPadY = 5f;

        void DrawTableRow(string[] rowCells, int rowIndex, bool isHeader)
        {
            var tf = isHeader ? BoldTypeface.Value : RegularTypeface.Value;
            using var font = new SKFont(tf, fontSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
            using var shaper = new SKShaper(tf);

            // Wrap each cell's lines first so we know the exact row height
            var wrappedCells = new List<List<List<(string Word, float Width)>>>(dataCols);
            int maxLines = 1;

            for (int c = 0; c < dataCols; c++)
            {
                string cellText = c < rowCells.Length ? rowCells[c] : string.Empty;
                float maxTextW = Math.Max(14f, dataColWidths[c] - (cellPadX * 2f));
                var lines = WrapTextIntoVisualLines(cellText, shaper, font, fontSize, maxTextW);
                if (lines.Count > maxLines)
                {
                    maxLines = lines.Count;
                }
                wrappedCells.Add(lines);
            }

            float rowHeight = Math.Max(fontSize + (cellPadY * 2f) + 4f, (maxLines * lineHeight) + (cellPadY * 2f));

            // Check if row fits on current page; if not, start a new page and repeat header if enabled
            if (ctx.CurrentY + rowHeight > ctx.PageBottom)
            {
                ctx.StartNewPage();
                if (!isHeader && options.RepeatHeaderRow && sheet.Rows.Count > 1)
                {
                    DrawTableRow(sheet.Rows[0], 0, isHeader: true);
                }
            }

            float y = ctx.CurrentY;

            // Background & border paints
            using var bgPaint = new SKPaint { Style = SKPaintStyle.Fill, IsAntialias = true };
            if (isHeader)
            {
                bgPaint.Color = accentColor;
            }
            else if (options.ZebraStriping && rowIndex % 2 == 0)
            {
                bgPaint.Color = new SKColor(248, 250, 252);
            }
            else
            {
                bgPaint.Color = SKColors.White;
            }

            using var borderPaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = isHeader ? accentColor.WithAlpha(220) : new SKColor(203, 213, 225),
                StrokeWidth = 0.7f,
                IsAntialias = true
            };

            using var textPaint = new SKPaint
            {
                IsAntialias = true,
                Color = isHeader ? SKColors.White : new SKColor(15, 23, 42)
            };

            float spaceW = font.MeasureText(" ");

            // Optional Row Number column
            if (showRowNum)
            {
                float rnX = isRtl ? (ctx.MarginLeft + ctx.ContentWidth - rowNumColWidth) : ctx.MarginLeft;
                var rnRect = new SKRect(rnX, y, rnX + rowNumColWidth, y + rowHeight);
                ctx.Canvas.DrawRect(rnRect, bgPaint);
                if (options.ShowGridlines || isHeader)
                {
                    ctx.Canvas.DrawRect(rnRect, borderPaint);
                }

                string rnLabel = isHeader ? "#" : rowIndex.ToString(CultureInfo.InvariantCulture);
                var rnWords = new List<(string Word, float Width)> { (rnLabel, MeasureWordWithEmojis(rnLabel, shaper, font, fontSize)) };
                DrawVisualLine(ctx.Canvas, rnWords, shaper, font, textPaint, fontSize, isRtl: false, rnX + 4f, rowNumColWidth - 8f, spaceW, y + cellPadY + fontSize);
            }

            // Draw data cells
            float currentX = isRtl
                ? (ctx.MarginLeft + ctx.ContentWidth - rowNumColWidth)
                : (ctx.MarginLeft + rowNumColWidth);

            for (int c = 0; c < dataCols; c++)
            {
                float colW = dataColWidths[c];
                float cellLeft = isRtl ? (currentX - colW) : currentX;
                var cellRect = new SKRect(cellLeft, y, cellLeft + colW, y + rowHeight);

                ctx.Canvas.DrawRect(cellRect, bgPaint);
                if (options.ShowGridlines || isHeader)
                {
                    ctx.Canvas.DrawRect(cellRect, borderPaint);
                }

                var cellLines = wrappedCells[c];
                string rawCellText = c < rowCells.Length ? rowCells[c] : string.Empty;
                bool cellRtl = isRtl || ContainsArabic(rawCellText);

                for (int lineIdx = 0; lineIdx < cellLines.Count; lineIdx++)
                {
                    float baselineY = y + cellPadY + fontSize + (lineIdx * lineHeight);
                    DrawVisualLine(
                        ctx.Canvas,
                        cellLines[lineIdx],
                        shaper,
                        font,
                        textPaint,
                        fontSize,
                        cellRtl,
                        cellLeft + cellPadX,
                        colW - (cellPadX * 2f),
                        spaceW,
                        baselineY);
                }

                currentX = isRtl ? (currentX - colW) : (currentX + colW);
            }

            ctx.AdvanceY(rowHeight);
        }

        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            bool isHeaderRow = r == 0;
            DrawTableRow(sheet.Rows[r], r, isHeaderRow);
        }

        ctx.AdvanceY(14f);
    }

    private void DrawSheetTitleBanner(PaginatedLayoutContext ctx, string sheetName, int rowCount, int colCount, SKColor accentColor, bool isRtl)
    {
        float bannerH = 26f;
        float y = ctx.CurrentY;
        var rect = new SKRect(ctx.MarginLeft, y, ctx.MarginLeft + ctx.ContentWidth, y + bannerH);

        using var bgPaint = new SKPaint
        {
            Style = SKPaintStyle.Fill,
            Color = new SKColor(241, 245, 249),
            IsAntialias = true
        };
        ctx.Canvas.DrawRoundRect(rect, 5f, 5f, bgPaint);

        // Accent bar on leading edge
        using var accentPaint = new SKPaint
        {
            Style = SKPaintStyle.Fill,
            Color = accentColor,
            IsAntialias = true
        };
        var barRect = isRtl
            ? new SKRect(rect.Right - 4f, y, rect.Right, y + bannerH)
            : new SKRect(rect.Left, y, rect.Left + 4f, y + bannerH);
        ctx.Canvas.DrawRect(barRect, accentPaint);

        using var font = new SKFont(BoldTypeface.Value, 11f) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
        using var shaper = new SKShaper(BoldTypeface.Value);
        using var textPaint = new SKPaint { IsAntialias = true, Color = new SKColor(15, 23, 42) };

        string label = $"{sheetName}  ({rowCount} × {colCount})";
        var words = label.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(w => (w, MeasureWordWithEmojis(w, shaper, font, 11f)))
            .ToList();

        DrawVisualLine(
            ctx.Canvas,
            words,
            shaper,
            font,
            textPaint,
            11f,
            isRtl || ContainsArabic(sheetName),
            ctx.MarginLeft + 10f,
            ctx.ContentWidth - 20f,
            font.MeasureText(" "),
            y + 17.5f);

        ctx.AdvanceY(bannerH + 6f);
    }

    private static float[] ComputeProportionalColumnWidths(List<string[]> rows, int colCount, float totalWidth)
    {
        var weights = new float[colCount];
        int sampleLimit = Math.Min(rows.Count, 60);

        for (int c = 0; c < colCount; c++)
        {
            int maxLen = 4;
            for (int r = 0; r < sampleLimit; r++)
            {
                if (c < rows[r].Length && rows[r][c] != null)
                {
                    int len = rows[r][c].Length;
                    if (len > maxLen) maxLen = len;
                }
            }
            // Clamp weight so neither tiny nor huge columns dominate excessively
            weights[c] = Math.Clamp((float)Math.Sqrt(maxLen), 1.8f, 6.5f);
        }

        float sumWeights = weights.Sum();
        var widths = new float[colCount];
        for (int c = 0; c < colCount; c++)
        {
            widths[c] = (weights[c] / sumWeights) * totalWidth;
        }

        return widths;
    }

    // =========================================================================
    // 3. Markdown AST Parser & Paginated Document Renderer
    // =========================================================================

    private enum MarkdownBlockType
    {
        Heading,
        Paragraph,
        UnorderedListItem,
        OrderedListItem,
        TaskListItem,
        Blockquote,
        CodeBlock,
        Table,
        HorizontalRule
    }

    private sealed class MarkdownBlock
    {
        public MarkdownBlockType Type { get; set; }
        public string Text { get; set; } = string.Empty;
        public int Level { get; set; } = 1; // Heading 1..6 or list nesting level
        public int OrderNumber { get; set; } = 1;
        public bool IsChecked { get; set; }
        public string CodeLanguage { get; set; } = string.Empty;
        public List<string[]> TableRows { get; set; } = new();
    }

    private sealed class InlineRun
    {
        public string Text { get; set; } = string.Empty;
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsCode { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsLink { get; set; }
    }

    private static List<MarkdownBlock> ParseMarkdownBlocks(string markdown)
    {
        var blocks = new List<MarkdownBlock>();
        var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

        int i = 0;
        while (i < lines.Length)
        {
            string rawLine = lines[i];
            string trimmed = rawLine.Trim();

            if (string.IsNullOrWhiteSpace(trimmed))
            {
                i++;
                continue;
            }

            // 1. Fenced Code Block (``` or ~~~)
            if (trimmed.StartsWith("```") || trimmed.StartsWith("~~~"))
            {
                string fence = trimmed.StartsWith("```") ? "```" : "~~~";
                string lang = trimmed[3..].Trim();
                var codeLines = new List<string>();
                i++;
                while (i < lines.Length && !lines[i].TrimStart().StartsWith(fence))
                {
                    codeLines.Add(lines[i]);
                    i++;
                }
                if (i < lines.Length) i++; // Skip closing fence

                blocks.Add(new MarkdownBlock
                {
                    Type = MarkdownBlockType.CodeBlock,
                    CodeLanguage = lang,
                    Text = string.Join('\n', codeLines)
                });
                continue;
            }

            // 2. Horizontal Rule (---, ***, ___)
            if (Regex.IsMatch(trimmed, @"^(\*{3,}|-{3,}|_{3,})$"))
            {
                blocks.Add(new MarkdownBlock { Type = MarkdownBlockType.HorizontalRule });
                i++;
                continue;
            }

            // 3. ATX Heading (# .. ######)
            var headingMatch = Regex.Match(trimmed, @"^(#{1,6})\s+(.+)$");
            if (headingMatch.Success)
            {
                blocks.Add(new MarkdownBlock
                {
                    Type = MarkdownBlockType.Heading,
                    Level = headingMatch.Groups[1].Value.Length,
                    Text = headingMatch.Groups[2].Value.TrimEnd('#').Trim()
                });
                i++;
                continue;
            }

            // 4. Pipe Table (| A | B |)
            if (trimmed.StartsWith('|') && trimmed.EndsWith('|') && trimmed.Length > 2)
            {
                var tableRows = new List<string[]>();
                while (i < lines.Length)
                {
                    string tLine = lines[i].Trim();
                    if (!tLine.StartsWith('|'))
                        break;

                    // Check if separator row (e.g. | --- | :---: |)
                    string inner = tLine.Trim('|');
                    var cells = inner.Split('|').Select(c => c.Trim()).ToArray();
                    bool isSeparator = cells.All(c => Regex.IsMatch(c, @"^:?-{2,}:?$"));
                    if (!isSeparator)
                    {
                        tableRows.Add(cells);
                    }
                    i++;
                }

                if (tableRows.Count > 0)
                {
                    blocks.Add(new MarkdownBlock
                    {
                        Type = MarkdownBlockType.Table,
                        TableRows = tableRows
                    });
                }
                continue;
            }

            // 5. Blockquote (> ...)
            if (trimmed.StartsWith('>'))
            {
                var quoteLines = new List<string>();
                while (i < lines.Length && lines[i].TrimStart().StartsWith('>'))
                {
                    string q = Regex.Replace(lines[i].TrimStart(), @"^>+\s?", "");
                    quoteLines.Add(q);
                    i++;
                }

                blocks.Add(new MarkdownBlock
                {
                    Type = MarkdownBlockType.Blockquote,
                    Text = string.Join(' ', quoteLines.Where(l => !string.IsNullOrWhiteSpace(l)))
                });
                continue;
            }

            // 6. Task List Item (- [ ] or - [x])
            var taskMatch = Regex.Match(rawLine, @"^(\s*)[-*+]\s+\[([ xX])\]\s+(.+)$");
            if (taskMatch.Success)
            {
                int indent = taskMatch.Groups[1].Value.Replace("\t", "  ").Length / 2;
                bool isChecked = !taskMatch.Groups[2].Value.Equals(" ", StringComparison.Ordinal);
                blocks.Add(new MarkdownBlock
                {
                    Type = MarkdownBlockType.TaskListItem,
                    Level = Math.Clamp(indent, 0, 4),
                    IsChecked = isChecked,
                    Text = taskMatch.Groups[3].Value.Trim()
                });
                i++;
                continue;
            }

            // 7. Unordered List Item (- , * , + )
            var ulMatch = Regex.Match(rawLine, @"^(\s*)[-*+]\s+(.+)$");
            if (ulMatch.Success)
            {
                int indent = ulMatch.Groups[1].Value.Replace("\t", "  ").Length / 2;
                blocks.Add(new MarkdownBlock
                {
                    Type = MarkdownBlockType.UnorderedListItem,
                    Level = Math.Clamp(indent, 0, 4),
                    Text = ulMatch.Groups[2].Value.Trim()
                });
                i++;
                continue;
            }

            // 8. Ordered List Item (1. , 2. )
            var olMatch = Regex.Match(rawLine, @"^(\s*)(\d+)[\.\)]\s+(.+)$");
            if (olMatch.Success)
            {
                int indent = olMatch.Groups[1].Value.Replace("\t", "  ").Length / 2;
                int.TryParse(olMatch.Groups[2].Value, out int num);
                blocks.Add(new MarkdownBlock
                {
                    Type = MarkdownBlockType.OrderedListItem,
                    Level = Math.Clamp(indent, 0, 4),
                    OrderNumber = num > 0 ? num : 1,
                    Text = olMatch.Groups[3].Value.Trim()
                });
                i++;
                continue;
            }

            // 9. Standard Paragraph (merge consecutive plain lines)
            var paraLines = new List<string> { trimmed };
            i++;
            while (i < lines.Length)
            {
                string nextTrim = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(nextTrim) ||
                    nextTrim.StartsWith('#') ||
                    nextTrim.StartsWith("```") ||
                    nextTrim.StartsWith("~~~") ||
                    nextTrim.StartsWith('>') ||
                    (nextTrim.StartsWith('|') && nextTrim.EndsWith('|')) ||
                    Regex.IsMatch(nextTrim, @"^(\*{3,}|-{3,}|_{3,})$") ||
                    Regex.IsMatch(lines[i], @"^\s*([-*+]|\d+[\.\)])\s+"))
                {
                    break;
                }
                paraLines.Add(nextTrim);
                i++;
            }

            blocks.Add(new MarkdownBlock
            {
                Type = MarkdownBlockType.Paragraph,
                Text = string.Join(' ', paraLines)
            });
        }

        return blocks;
    }

    private void RenderMarkdownBlock(
        PaginatedLayoutContext ctx,
        MarkdownBlock block,
        float scale,
        SKColor accentColor,
        SpreadsheetMarkdownOptions options)
    {
        switch (block.Type)
        {
            case MarkdownBlockType.HorizontalRule:
            {
                ctx.EnsureSpace(20f);
                ctx.AdvanceY(8f);
                using var hrPaint = new SKPaint
                {
                    Style = SKPaintStyle.Stroke,
                    Color = new SKColor(203, 213, 225),
                    StrokeWidth = 1.2f,
                    IsAntialias = true
                };
                ctx.Canvas.DrawLine(ctx.MarginLeft, ctx.CurrentY, ctx.MarginLeft + ctx.ContentWidth, ctx.CurrentY, hrPaint);
                ctx.AdvanceY(12f);
                break;
            }

            case MarkdownBlockType.Heading:
            {
                float baseSize = block.Level switch
                {
                    1 => 20f,
                    2 => 16.5f,
                    3 => 14f,
                    4 => 12.5f,
                    _ => 11.5f
                };
                float fontSize = baseSize * scale;
                float lineHeight = fontSize * 1.45f;

                ctx.EnsureSpace(lineHeight + 16f);
                if (ctx.CurrentY > ctx.MarginTop + 6f)
                {
                    ctx.AdvanceY(block.Level <= 2 ? 8f : 4f);
                }

                SKColor headingColor = block.Level == 1
                    ? accentColor
                    : new SKColor(15, 23, 42);

                bool isRtl = ContainsArabic(block.Text);
                RenderFormattedParagraph(ctx, block.Text, fontSize, lineHeight, headingColor, forceBold: true, isRtl, ctx.MarginLeft, ctx.ContentWidth);

                if (block.Level <= 2)
                {
                    ctx.AdvanceY(3f);
                    using var rulePaint = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = block.Level == 1 ? accentColor.WithAlpha(90) : new SKColor(226, 232, 240),
                        StrokeWidth = block.Level == 1 ? 1.5f : 1f,
                        IsAntialias = true
                    };
                    ctx.Canvas.DrawLine(ctx.MarginLeft, ctx.CurrentY, ctx.MarginLeft + ctx.ContentWidth, ctx.CurrentY, rulePaint);
                    ctx.AdvanceY(7f);
                }
                else
                {
                    ctx.AdvanceY(4f);
                }
                break;
            }

            case MarkdownBlockType.Paragraph:
            {
                float fontSize = 11.2f * scale;
                float lineHeight = fontSize * 1.55f;
                bool isRtl = ContainsArabic(block.Text);
                RenderFormattedParagraph(ctx, block.Text, fontSize, lineHeight, new SKColor(30, 41, 59), forceBold: false, isRtl, ctx.MarginLeft, ctx.ContentWidth);
                ctx.AdvanceY(7f);
                break;
            }

            case MarkdownBlockType.UnorderedListItem:
            case MarkdownBlockType.OrderedListItem:
            case MarkdownBlockType.TaskListItem:
            {
                float fontSize = 11f * scale;
                float lineHeight = fontSize * 1.5f;
                bool isRtl = ContainsArabic(block.Text);
                float indentOffset = 16f + (block.Level * 16f);
                float textWidth = ctx.ContentWidth - indentOffset;
                float textStartX = isRtl ? ctx.MarginLeft : (ctx.MarginLeft + indentOffset);

                ctx.EnsureSpace(lineHeight);
                float markerY = ctx.CurrentY;

                // Draw bullet, number, or task checkbox
                if (block.Type == MarkdownBlockType.TaskListItem)
                {
                    float boxSize = fontSize * 0.95f;
                    float boxX = isRtl
                        ? (ctx.MarginLeft + ctx.ContentWidth - indentOffset + 3f)
                        : (ctx.MarginLeft + (block.Level * 16f) + 2f);
                    float boxY = markerY + (fontSize * 0.18f);
                    var boxRect = new SKRect(boxX, boxY, boxX + boxSize, boxY + boxSize);

                    if (block.IsChecked)
                    {
                        using var fillPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = new SKColor(5, 150, 105), IsAntialias = true };
                        ctx.Canvas.DrawRoundRect(boxRect, 2.5f, 2.5f, fillPaint);

                        using var checkPaint = new SKPaint
                        {
                            Style = SKPaintStyle.Stroke,
                            Color = SKColors.White,
                            StrokeWidth = 1.6f,
                            StrokeCap = SKStrokeCap.Round,
                            StrokeJoin = SKStrokeJoin.Round,
                            IsAntialias = true
                        };
                        ctx.Canvas.DrawLine(boxX + boxSize * 0.22f, boxY + boxSize * 0.52f, boxX + boxSize * 0.44f, boxY + boxSize * 0.74f, checkPaint);
                        ctx.Canvas.DrawLine(boxX + boxSize * 0.44f, boxY + boxSize * 0.74f, boxX + boxSize * 0.78f, boxY + boxSize * 0.26f, checkPaint);
                    }
                    else
                    {
                        using var borderPaint = new SKPaint
                        {
                            Style = SKPaintStyle.Stroke,
                            Color = new SKColor(148, 163, 184),
                            StrokeWidth = 1.2f,
                            IsAntialias = true
                        };
                        ctx.Canvas.DrawRoundRect(boxRect, 2.5f, 2.5f, borderPaint);
                    }
                }
                else
                {
                    string marker = block.Type == MarkdownBlockType.OrderedListItem
                        ? $"{block.OrderNumber}."
                        : "•";

                    using var mFont = new SKFont(BoldTypeface.Value, fontSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
                    using var mPaint = new SKPaint { IsAntialias = true, Color = accentColor };
                    float markerX = isRtl
                        ? (ctx.MarginLeft + ctx.ContentWidth - (block.Level * 16f) - mFont.MeasureText(marker) - 2f)
                        : (ctx.MarginLeft + (block.Level * 16f) + 2f);
                    ctx.Canvas.DrawText(marker, markerX, markerY + fontSize, SKTextAlign.Left, mFont, mPaint);
                }

                RenderFormattedParagraph(ctx, block.Text, fontSize, lineHeight, new SKColor(30, 41, 59), forceBold: false, isRtl, textStartX, textWidth);
                ctx.AdvanceY(3f);
                break;
            }

            case MarkdownBlockType.Blockquote:
            {
                float fontSize = 10.8f * scale;
                float lineHeight = fontSize * 1.5f;
                bool isRtl = ContainsArabic(block.Text);
                float padX = 14f;
                float padY = 8f;

                // Measure height by dry-wrapping
                using var font = new SKFont(ItalicTypeface.Value, fontSize);
                using var shaper = new SKShaper(ItalicTypeface.Value);
                var visualLines = WrapTextIntoVisualLines(StripMarkdownMarkers(block.Text), shaper, font, fontSize, ctx.ContentWidth - (padX * 2f));
                float boxHeight = (visualLines.Count * lineHeight) + (padY * 2f);

                ctx.EnsureSpace(boxHeight + 8f);
                float startY = ctx.CurrentY;
                var boxRect = new SKRect(ctx.MarginLeft, startY, ctx.MarginLeft + ctx.ContentWidth, startY + boxHeight);

                using var bgPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = new SKColor(248, 250, 252), IsAntialias = true };
                ctx.Canvas.DrawRoundRect(boxRect, 4f, 4f, bgPaint);

                using var barPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = accentColor, IsAntialias = true };
                var barRect = isRtl
                    ? new SKRect(boxRect.Right - 3.5f, startY, boxRect.Right, startY + boxHeight)
                    : new SKRect(boxRect.Left, startY, boxRect.Left + 3.5f, startY + boxHeight);
                ctx.Canvas.DrawRect(barRect, barPaint);

                ctx.AdvanceY(padY);
                RenderFormattedParagraph(ctx, block.Text, fontSize, lineHeight, new SKColor(71, 85, 105), forceBold: false, isRtl, ctx.MarginLeft + padX, ctx.ContentWidth - (padX * 2f));
                ctx.AdvanceY(padY + 6f);
                break;
            }

            case MarkdownBlockType.CodeBlock:
            {
                float fontSize = 9.5f * scale;
                float lineHeight = fontSize * 1.45f;
                var codeLines = block.Text.Split('\n');
                float padX = 12f;
                float padY = 10f;
                float boxHeight = (codeLines.Length * lineHeight) + (padY * 2f);

                ctx.EnsureSpace(Math.Min(boxHeight + 10f, ctx.ContentHeight * 0.8f));
                float startY = ctx.CurrentY;

                // Render code block background
                float actualBoxHeight = Math.Min(boxHeight, ctx.PageBottom - startY);
                var codeRect = new SKRect(ctx.MarginLeft, startY, ctx.MarginLeft + ctx.ContentWidth, startY + actualBoxHeight);

                using var bgPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = new SKColor(241, 245, 249), IsAntialias = true };
                using var borderPaint = new SKPaint { Style = SKPaintStyle.Stroke, Color = new SKColor(203, 213, 225), StrokeWidth = 0.8f, IsAntialias = true };
                ctx.Canvas.DrawRoundRect(codeRect, 6f, 6f, bgPaint);
                ctx.Canvas.DrawRoundRect(codeRect, 6f, 6f, borderPaint);

                ctx.AdvanceY(padY);

                using var monoFont = new SKFont(MonoTypeface.Value, fontSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
                using var monoShaper = new SKShaper(MonoTypeface.Value);
                using var codeTextPaint = new SKPaint { IsAntialias = true, Color = new SKColor(15, 23, 42) };

                foreach (var cLine in codeLines)
                {
                    ctx.EnsureSpace(lineHeight);
                    string displayLine = string.IsNullOrEmpty(cLine) ? " " : cLine.Replace("\t", "    ");
                    bool lineRtl = ContainsArabic(displayLine);
                    ctx.Canvas.DrawShapedText(monoShaper, displayLine, ctx.MarginLeft + padX, ctx.CurrentY + fontSize, SKTextAlign.Left, monoFont, codeTextPaint);
                    ctx.AdvanceY(lineHeight);
                }

                ctx.AdvanceY(padY + 6f);
                break;
            }

            case MarkdownBlockType.Table:
            {
                if (block.TableRows.Count == 0)
                    break;

                int maxCols = block.TableRows.Max(r => r.Length);
                if (maxCols <= 0)
                    break;

                // Normalize rows to maxCols and strip markdown inline markers inside cells
                var normalized = block.TableRows
                    .Select(r =>
                    {
                        var arr = new string[maxCols];
                        for (int c = 0; c < maxCols; c++)
                        {
                            arr[c] = c < r.Length ? StripMarkdownMarkers(r[c]) : string.Empty;
                        }
                        return arr;
                    })
                    .ToList();

                bool tableRtl = normalized.SelectMany(r => r).Any(ContainsArabic);
                var syntheticSheet = new ExtractedWorksheet
                {
                    Name = string.Empty,
                    Rows = normalized,
                    ColumnCount = maxCols,
                    IsRtl = tableRtl
                };

                var tableOpts = new SpreadsheetMarkdownOptions
                {
                    ShowSheetTitle = false,
                    ShowRowNumbers = false,
                    RepeatHeaderRow = true,
                    ZebraStriping = true,
                    ShowGridlines = true
                };

                ctx.AdvanceY(4f);
                RenderExcelSheet(ctx, syntheticSheet, tableOpts, accentColor, tableRtl);
                break;
            }
        }
    }

    // =========================================================================
    // 4. Inline Markdown Formatting & Word-Level BiDi / Emoji Shaper
    // =========================================================================

    private static List<InlineRun> ParseInlineRuns(string text, bool defaultBold = false)
    {
        var runs = new List<InlineRun>();
        if (string.IsNullOrEmpty(text))
            return runs;

        // Pattern matches: `code`, **bold**, __bold__, *italic*, _italic_, ~~strike~~, [link](url)
        var pattern = new Regex(@"(`[^`]+`|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|__[^_]+__|~~[^~]+~~|\*[^*]+\*|_[^_]+_|\[[^\]]+\]\([^)]+\))");
        int lastIdx = 0;

        foreach (Match match in pattern.Matches(text))
        {
            if (match.Index > lastIdx)
            {
                runs.Add(new InlineRun
                {
                    Text = text[lastIdx..match.Index],
                    IsBold = defaultBold
                });
            }

            string token = match.Value;
            if (token.StartsWith('`') && token.EndsWith('`') && token.Length > 2)
            {
                runs.Add(new InlineRun { Text = token[1..^1], IsCode = true, IsBold = defaultBold });
            }
            else if (token.StartsWith("***") && token.EndsWith("***") && token.Length > 6)
            {
                runs.Add(new InlineRun { Text = token[3..^3], IsBold = true, IsItalic = true });
            }
            else if ((token.StartsWith("**") && token.EndsWith("**")) || (token.StartsWith("__") && token.EndsWith("__")))
            {
                runs.Add(new InlineRun { Text = token[2..^2], IsBold = true });
            }
            else if (token.StartsWith("~~") && token.EndsWith("~~") && token.Length > 4)
            {
                runs.Add(new InlineRun { Text = token[2..^2], IsStrikethrough = true, IsBold = defaultBold });
            }
            else if ((token.StartsWith('*') && token.EndsWith('*')) || (token.StartsWith('_') && token.EndsWith('_')))
            {
                runs.Add(new InlineRun { Text = token[1..^1], IsItalic = true, IsBold = defaultBold });
            }
            else if (token.StartsWith('['))
            {
                int closeBracket = token.IndexOf("](", StringComparison.Ordinal);
                if (closeBracket > 1)
                {
                    string linkLabel = token[1..closeBracket];
                    runs.Add(new InlineRun { Text = linkLabel, IsLink = true, IsBold = defaultBold });
                }
            }

            lastIdx = match.Index + match.Length;
        }

        if (lastIdx < text.Length)
        {
            runs.Add(new InlineRun
            {
                Text = text[lastIdx..],
                IsBold = defaultBold
            });
        }

        return runs;
    }

    private static string StripMarkdownMarkers(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var runs = ParseInlineRuns(text);
        return string.Concat(runs.Select(r => r.Text));
    }

    private sealed class StyledWord
    {
        public string Word { get; set; } = string.Empty;
        public float Width { get; set; }
        public SKTypeface Typeface { get; set; } = null!;
        public bool IsCode { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsLink { get; set; }
    }

    private void RenderFormattedParagraph(
        PaginatedLayoutContext ctx,
        string rawText,
        float fontSize,
        float lineHeight,
        SKColor baseColor,
        bool forceBold,
        bool isRtl,
        float startX,
        float maxWidth)
    {
        var runs = ParseInlineRuns(rawText, forceBold);
        var styledWords = new List<StyledWord>();

        foreach (var run in runs)
        {
            var tf = run.IsCode
                ? MonoTypeface.Value
                : (run.IsBold && run.IsItalic)
                    ? BoldItalicTypeface.Value
                    : run.IsBold
                        ? BoldTypeface.Value
                        : run.IsItalic
                            ? ItalicTypeface.Value
                            : RegularTypeface.Value;

            float effectiveFontSize = run.IsCode ? fontSize * 0.92f : fontSize;
            using var font = new SKFont(tf, effectiveFontSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
            using var shaper = new SKShaper(tf);

            var words = run.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            foreach (var w in words)
            {
                float wWidth = MeasureWordWithEmojis(w, shaper, font, effectiveFontSize);
                if (run.IsCode)
                {
                    wWidth += 6f; // Padding for code pill
                }

                styledWords.Add(new StyledWord
                {
                    Word = w,
                    Width = wWidth,
                    Typeface = tf,
                    IsCode = run.IsCode,
                    IsStrikethrough = run.IsStrikethrough,
                    IsLink = run.IsLink
                });
            }
        }

        if (styledWords.Count == 0)
            return;

        using var baseFont = new SKFont(RegularTypeface.Value, fontSize);
        float spaceWidth = baseFont.MeasureText(" ");

        var currentLine = new List<StyledWord>();
        float currentWidth = 0f;

        foreach (var sw in styledWords)
        {
            float projected = currentLine.Count == 0 ? sw.Width : currentWidth + spaceWidth + sw.Width;
            if (projected > maxWidth && currentLine.Count > 0)
            {
                ctx.EnsureSpace(lineHeight);
                DrawStyledVisualLine(ctx.Canvas, currentLine, fontSize, baseColor, isRtl, startX, maxWidth, spaceWidth, ctx.CurrentY + fontSize);
                ctx.AdvanceY(lineHeight);

                currentLine.Clear();
                currentLine.Add(sw);
                currentWidth = sw.Width;
            }
            else
            {
                currentLine.Add(sw);
                currentWidth = projected;
            }
        }

        if (currentLine.Count > 0)
        {
            ctx.EnsureSpace(lineHeight);
            DrawStyledVisualLine(ctx.Canvas, currentLine, fontSize, baseColor, isRtl, startX, maxWidth, spaceWidth, ctx.CurrentY + fontSize);
            ctx.AdvanceY(lineHeight);
        }
    }

    private static void DrawStyledVisualLine(
        SKCanvas canvas,
        List<StyledWord> words,
        float fontSize,
        SKColor baseColor,
        bool isRtl,
        float startX,
        float maxWidth,
        float spaceWidth,
        float baselineY)
    {
        float cursor = isRtl ? (startX + maxWidth) : startX;

        foreach (var sw in words)
        {
            float wordLeft = isRtl ? (cursor - sw.Width) : cursor;
            float effectiveSize = sw.IsCode ? fontSize * 0.92f : fontSize;

            using var font = new SKFont(sw.Typeface, effectiveSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
            using var shaper = new SKShaper(sw.Typeface);
            using var paint = new SKPaint
            {
                IsAntialias = true,
                Color = sw.IsLink ? new SKColor(37, 99, 235) : baseColor
            };

            float drawX = wordLeft;
            if (sw.IsCode)
            {
                using var pillPaint = new SKPaint
                {
                    Style = SKPaintStyle.Fill,
                    Color = new SKColor(241, 245, 249),
                    IsAntialias = true
                };
                var pillRect = new SKRect(wordLeft, baselineY - (fontSize * 0.88f), wordLeft + sw.Width, baselineY + (fontSize * 0.22f));
                canvas.DrawRoundRect(pillRect, 3f, 3f, pillPaint);
                drawX += 3f;
            }

            DrawWordTokens(canvas, sw.Word, shaper, font, paint, effectiveSize, isRtl, drawX, sw.IsCode ? sw.Width - 6f : sw.Width, baselineY);

            if (sw.IsStrikethrough)
            {
                using var strikePaint = new SKPaint
                {
                    Style = SKPaintStyle.Stroke,
                    Color = paint.Color,
                    StrokeWidth = 1f,
                    IsAntialias = true
                };
                float strikeY = baselineY - (fontSize * 0.32f);
                canvas.DrawLine(wordLeft, strikeY, wordLeft + sw.Width, strikeY, strikePaint);
            }
            else if (sw.IsLink)
            {
                using var linkLinePaint = new SKPaint
                {
                    Style = SKPaintStyle.Stroke,
                    Color = new SKColor(37, 99, 235, 160),
                    StrokeWidth = 0.8f,
                    IsAntialias = true
                };
                canvas.DrawLine(wordLeft, baselineY + 1.5f, wordLeft + sw.Width, baselineY + 1.5f, linkLinePaint);
            }

            cursor = isRtl ? (wordLeft - spaceWidth) : (wordLeft + sw.Width + spaceWidth);
        }
    }

    private static List<List<(string Word, float Width)>> WrapTextIntoVisualLines(
        string text,
        SKShaper shaper,
        SKFont font,
        float fontSize,
        float maxWidth)
    {
        var result = new List<List<(string Word, float Width)>>();
        if (string.IsNullOrWhiteSpace(text))
            return result;

        float spaceW = font.MeasureText(" ");
        var logicalLines = text.Replace("\r\n", "\n").Split('\n');

        foreach (var logical in logicalLines)
        {
            var words = logical.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0)
                continue;

            var currentLine = new List<(string Word, float Width)>();
            float currentWidth = 0f;

            foreach (var word in words)
            {
                float wWidth = MeasureWordWithEmojis(word, shaper, font, fontSize);
                float projected = currentLine.Count == 0 ? wWidth : currentWidth + spaceW + wWidth;

                if (projected > maxWidth && currentLine.Count > 0)
                {
                    result.Add(currentLine);
                    currentLine = new List<(string Word, float Width)> { (word, wWidth) };
                    currentWidth = wWidth;
                }
                else
                {
                    currentLine.Add((word, wWidth));
                    currentWidth = projected;
                }
            }

            if (currentLine.Count > 0)
            {
                result.Add(currentLine);
            }
        }

        return result;
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

    private static (float Width, float Height) GetPageDimensions(PdfPageSizePreset preset, PdfPageOrientation orientation)
    {
        var (w, h) = preset switch
        {
            PdfPageSizePreset.Letter => (612f, 792f),
            _ => (595f, 842f) // A4
        };

        return orientation == PdfPageOrientation.Landscape ? (h, w) : (w, h);
    }

    private static SKColor ParseColorOrDefault(string? hex, SKColor fallback)
    {
        if (!string.IsNullOrWhiteSpace(hex) && SKColor.TryParse(hex.Trim(), out var parsed))
        {
            return parsed;
        }
        return fallback;
    }

    public static bool ContainsArabic(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

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

    // =========================================================================
    // 5. Paginated Skia Layout Context
    // =========================================================================

    private sealed class PaginatedLayoutContext : IDisposable
    {
        private readonly SKDocument _pdfDoc;
        private readonly string _docTitle;
        private readonly bool _includePageNumbers;
        private bool _pageOpen;

        public float PageWidth { get; }
        public float PageHeight { get; }
        public float MarginLeft { get; }
        public float MarginTop { get; }
        public float MarginBottom { get; }
        public float ContentWidth { get; }
        public float ContentHeight { get; }
        public float PageBottom => PageHeight - MarginBottom;

        public SKCanvas Canvas { get; private set; } = null!;
        public float CurrentY { get; private set; }
        public int PageCount { get; private set; }

        public PaginatedLayoutContext(
            SKDocument pdfDoc,
            float pageWidth,
            float pageHeight,
            float marginH,
            float marginV,
            string docTitle,
            bool includePageNumbers)
        {
            _pdfDoc = pdfDoc;
            PageWidth = pageWidth;
            PageHeight = pageHeight;
            MarginLeft = marginH;
            MarginTop = marginV;
            MarginBottom = includePageNumbers ? marginV + 14f : marginV;
            ContentWidth = pageWidth - (marginH * 2f);
            ContentHeight = PageBottom - MarginTop;
            _docTitle = docTitle;
            _includePageNumbers = includePageNumbers;

            StartNewPage();
        }

        public void StartNewPage()
        {
            if (_pageOpen)
            {
                DrawFooterIfEnabled();
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
            if (CurrentY + requiredHeight > PageBottom && CurrentY > MarginTop + 1f)
            {
                StartNewPage();
            }
        }

        public void Finish()
        {
            if (_pageOpen)
            {
                DrawFooterIfEnabled();
                _pdfDoc.EndPage();
                _pageOpen = false;
            }
        }

        private void DrawFooterIfEnabled()
        {
            if (!_includePageNumbers)
                return;

            float footerY = PageHeight - 18f;
            using var font = new SKFont(RegularTypeface.Value, 8.5f) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
            using var shaper = new SKShaper(RegularTypeface.Value);
            using var paint = new SKPaint { IsAntialias = true, Color = new SKColor(148, 163, 184) };

            string pageLabel = $"Page {PageCount}";
            float labelW = font.MeasureText(pageLabel);
            Canvas.DrawText(pageLabel, PageWidth - MarginLeft - labelW, footerY, SKTextAlign.Left, font, paint);

            if (!string.IsNullOrWhiteSpace(_docTitle))
            {
                Canvas.DrawShapedText(shaper, _docTitle, MarginLeft, footerY, SKTextAlign.Left, font, paint);
            }
        }

        public void Dispose()
        {
            Finish();
        }
    }
}
