namespace Mnshor.Models;

public enum ConversionInputMode
{
    Excel,
    Markdown
}

public enum PdfPageOrientation
{
    Portrait,
    Landscape
}

public enum PdfPageSizePreset
{
    A4,
    Letter
}

public enum TableDirectionMode
{
    Auto,
    Ltr,
    Rtl
}

public enum MarkdownFontSizeScale
{
    Small,
    Medium,
    Large
}

public class SpreadsheetMarkdownOptions
{
    public ConversionInputMode Mode { get; set; } = ConversionInputMode.Excel;

    public PdfPageOrientation Orientation { get; set; } = PdfPageOrientation.Landscape;

    public PdfPageSizePreset PageSize { get; set; } = PdfPageSizePreset.A4;

    // Excel-specific options
    public HashSet<string> SelectedSheetNames { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public bool RepeatHeaderRow { get; set; } = true;

    public bool ZebraStriping { get; set; } = true;

    public bool ShowGridlines { get; set; } = true;

    public bool ShowSheetTitle { get; set; } = true;

    public bool ShowRowNumbers { get; set; } = false;

    public TableDirectionMode DirectionMode { get; set; } = TableDirectionMode.Auto;

    public string HeaderAccentHex { get; set; } = "#1E293B";

    // Markdown-specific options
    public MarkdownFontSizeScale FontSizeScale { get; set; } = MarkdownFontSizeScale.Medium;

    public bool IncludePageNumbers { get; set; } = true;
}

public class ExcelWorkbookInfo
{
    public bool IsValid { get; set; }
    public string? ErrorMessage { get; set; }
    public List<ExcelSheetPreview> Sheets { get; set; } = new();
}

public class ExcelSheetPreview
{
    public string Name { get; set; } = string.Empty;
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }
    public bool IsRtlDetected { get; set; }
    public List<string[]> PreviewRows { get; set; } = new();
}
