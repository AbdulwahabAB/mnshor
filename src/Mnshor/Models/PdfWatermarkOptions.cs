namespace Mnshor.Models;

public enum WatermarkMode
{
    Text,
    Image
}

public enum WatermarkPosition
{
    TopLeft,
    TopCenter,
    TopRight,
    MiddleLeft,
    Center,
    MiddleRight,
    BottomLeft,
    BottomCenter,
    BottomRight
}

public enum WatermarkLayer
{
    OverContent,
    BelowContent
}

public class PdfWatermarkOptions
{
    public WatermarkMode Mode { get; set; } = WatermarkMode.Text;

    // Text Watermark Properties
    public string Text { get; set; } = "CONFIDENTIAL";
    public string FontFamily { get; set; } = "Segoe UI";
    public float FontSize { get; set; } = 48f;
    public bool IsBold { get; set; } = true;
    public bool IsItalic { get; set; } = false;
    public string TextColorHex { get; set; } = "#EF4444";

    // Image Watermark Properties
    public byte[]? ImageData { get; set; }
    public string? ImageFileName { get; set; }
    public string? ImageDataUrl { get; set; }
    public int ImageScalePercent { get; set; } = 35;

    // Position & Tiling
    public WatermarkPosition Position { get; set; } = WatermarkPosition.Center;
    public bool IsMosaic { get; set; } = false;

    // Appearance & Layer
    public int OpacityPercent { get; set; } = 50;
    public int RotationDegrees { get; set; } = -45;
    public WatermarkLayer Layer { get; set; } = WatermarkLayer.OverContent;

    // Page Targeting
    public string PageScope { get; set; } = "all"; // "all" or "range"
    public string? CustomRange { get; set; }
}
