namespace Mnshor.Models;

public class ExecutionResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string OutputFileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = "application/pdf";
    public byte[] OutputData { get; set; } = Array.Empty<byte>();
    public long ElapsedMilliseconds { get; set; }
    public string Summary { get; set; } = string.Empty;

    public static ExecutionResult Fail(string error, long elapsedMs = 0) => new()
    {
        Success = false,
        ErrorMessage = error,
        ElapsedMilliseconds = elapsedMs
    };

    public static ExecutionResult Ok(byte[] data, string fileName, string summary, long elapsedMs, string contentType = "application/pdf") => new()
    {
        Success = true,
        OutputData = data,
        OutputFileName = fileName,
        Summary = summary,
        ElapsedMilliseconds = elapsedMs,
        ContentType = contentType
    };
}
