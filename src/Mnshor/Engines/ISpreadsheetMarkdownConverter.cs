using Mnshor.Models;

namespace Mnshor.Engines;

public interface ISpreadsheetMarkdownConverter
{
    ExcelWorkbookInfo InspectWorkbook(byte[] xlsxData);

    Task<ExecutionResult> ConvertExcelToPdfAsync(DocumentItem document, SpreadsheetMarkdownOptions? options = null);

    Task<ExecutionResult> ConvertMarkdownToPdfAsync(string markdownText, string documentTitle, SpreadsheetMarkdownOptions? options = null);
}
