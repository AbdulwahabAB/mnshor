namespace Mnshor.Models;

public class PdfMetadataModel
{
    public string Title { get; set; } = string.Empty;
    public string Author { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Keywords { get; set; } = string.Empty;
    public string Creator { get; set; } = string.Empty;
    public string Producer { get; set; } = string.Empty;
    public DateTime? CreationDate { get; set; }
    public DateTime? ModificationDate { get; set; }
    public string PdfVersion { get; set; } = string.Empty;

    public bool HasAnyMetadata =>
        !string.IsNullOrWhiteSpace(Title) ||
        !string.IsNullOrWhiteSpace(Author) ||
        !string.IsNullOrWhiteSpace(Subject) ||
        !string.IsNullOrWhiteSpace(Keywords) ||
        !string.IsNullOrWhiteSpace(Creator);

    public void ClearEditableFields()
    {
        Title = string.Empty;
        Author = string.Empty;
        Subject = string.Empty;
        Keywords = string.Empty;
        Creator = string.Empty;
    }
}
