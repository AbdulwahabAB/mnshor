using Mnshor.Models;

namespace Mnshor.Engines;

public interface IWordConverter
{
    Task<ExecutionResult> ConvertToPdfAsync(DocumentItem document);
}
